"""
downloader.py
─────────────
Scrapes CafeF's financial reports listing page for each company ticker,
filters for consolidated + audited + annual PDFs for target years,
then downloads them to data/raw/{TICKER}_{YEAR}.pdf.

CafeF's report table is JavaScript-rendered (loaded via AJAX after page load).
requests + BeautifulSoup only sees the bare HTML skeleton — no links.
Playwright is required to execute JavaScript and get the fully rendered DOM.
"""

import re
import time
import requests
from pathlib import Path
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout

from config import (
    DATA_RAW, DEFAULT_YEARS,
    CAFEF_URL_TEMPLATE, HTTP_HEADERS,
)


def _is_target_row(row_text: str, years: list[int]) -> tuple[bool, int | None]:
    """
    Filter by Vietnamese row text — far more reliable than URL matching.
    Row text example: "Báo cáo tài chính hợp nhất năm 2024 (đã kiểm toán) | CN/2024"

    Returns (is_match, year_or_None).
    """
    text = row_text.lower()

    # Must be consolidated (hợp nhất)
    if 'hợp nhất' not in text:
        return False, None

    # Must be audited (đã kiểm toán)
    if 'đã kiểm toán' not in text:
        return False, None

    # Must be annual — exclude quarterly (quý)
    if 'quý' in text:
        return False, None

    # Match target year
    for year in years:
        if str(year) in row_text:
            return True, year

    return False, None



def _get_pdf_links_playwright(ticker: str, years: list[int]) -> list[dict]:
    """
    Use Playwright to render the CafeF page and extract PDF links.
    Returns list of {ticker, year, url} dicts.
    """
    page_url = CAFEF_URL_TEMPLATE.format(ticker=ticker.lower())
    print(f"  Fetching (Playwright): {page_url}")

    results = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            user_agent=HTTP_HEADERS["User-Agent"],
            locale="vi-VN",
        )
        page = context.new_page()

        try:
            # Wait for network to settle (JS data loads via AJAX)
            page.goto(page_url, wait_until="networkidle", timeout=30000)
        except PlaywrightTimeout:
            print(f"  ⚠ Page load timeout — trying anyway with what loaded")

        # Give extra time for AJAX table to render
        page.wait_for_timeout(2000)

        # Find all anchor tags with .pdf in href
        pdf_links = page.query_selector_all("a[href*='.pdf'], a[href*='.PDF']")
        print(f"  Found {len(pdf_links)} PDF links in rendered page")

        seen_years = set()

        for link in pdf_links:
            href = link.get_attribute("href") or ""

            # Normalise to absolute URL
            if href.startswith("//"):
                href = "https:" + href
            elif href.startswith("/"):
                href = "https://cafef.vn" + href
            elif not href.startswith("http"):
                continue

            # Get row text for reliable Vietnamese-based filtering
            try:
                row_text = link.evaluate(
                    "el => el.closest('tr') ? el.closest('tr').innerText : ''"
                ).replace("\n", " ").strip()
            except Exception:
                row_text = ""

            is_match, year = _is_target_row(row_text, years)
            if is_match and year not in seen_years:
                results.append({"ticker": ticker, "year": year, "url": href})
                seen_years.add(year)
                print(f"    ✓ {ticker} {year}: {row_text[:60]}")
                print(f"       → {href[-70:]}")

        missing = [y for y in years if y not in seen_years]
        if missing:
            print(f"  ⚠ No PDF found for years: {missing}")

        browser.close()

    return results


_CDN_VARIANTS = ["cafefnew.mediacdn.vn", "cafef1.mediacdn.vn"]


def _cdn_fallbacks(url: str) -> list[str]:
    """
    Return all CDN variants of a mediacdn.vn URL to try in order.

    CafeF hosts files on either cafefnew.mediacdn.vn or cafef1.mediacdn.vn
    with no reliable rule for which one a given file lives on — it varies by
    year, company, and upload batch. The scraped URL may point to the wrong
    host, causing a 404. We build the full list of variants so _download_pdf
    can try each one until one succeeds.

    Non-mediacdn URLs (e.g. /BCTC/000... paths served from cafef.vn) are
    returned unchanged as a single-element list.
    """
    if "mediacdn.vn" not in url:
        return [url]

    variants = []
    for cdn in _CDN_VARIANTS:
        # Replace whichever mediacdn subdomain is currently in the URL
        candidate = re.sub(r"cafef\w+\.mediacdn\.vn", cdn, url)
        if candidate not in variants:
            variants.append(candidate)
    return variants


def _download_pdf(url: str, dest: Path) -> bool:
    """
    Download a PDF and save to dest. Returns True on success.

    For mediacdn.vn URLs, automatically retries with alternate CDN hostnames
    on 404 — CafeF distributes files across cafefnew / cafef1 unpredictably.
    """
    if dest.exists():
        print(f"  → Already exists, skipping: {dest.name}")
        return True

    candidates = _cdn_fallbacks(url)

    for attempt_url in candidates:
        try:
            resp = requests.get(attempt_url, headers=HTTP_HEADERS, timeout=60, stream=True)
            resp.raise_for_status()

            with open(dest, "wb") as f:
                for chunk in resp.iter_content(chunk_size=8192):
                    f.write(chunk)

            size_kb = dest.stat().st_size // 1024
            if attempt_url != url:
                print(f"  ✓ Downloaded {dest.name} ({size_kb} KB) [CDN fallback: {attempt_url.split('/')[2]}]")
            else:
                print(f"  ✓ Downloaded {dest.name} ({size_kb} KB)")
            return True

        except requests.HTTPError as e:
            if e.response is not None and e.response.status_code == 404 and len(candidates) > 1:
                print(f"  ⚠ 404 on {attempt_url.split('/')[2]} — trying next CDN...")
                if dest.exists():
                    dest.unlink()
                continue
            print(f"  ✗ Download failed for {dest.name}: {e}")
            if dest.exists():
                dest.unlink()
            return False

        except requests.RequestException as e:
            print(f"  ✗ Download failed for {dest.name}: {e}")
            if dest.exists():
                dest.unlink()
            return False

    print(f"  ✗ All CDN variants failed for {dest.name}")
    return False


def run(tickers: list[str] | None = None, years: list[int] | None = None) -> list[Path]:
    """
    Main entry point. Downloads all matching PDFs.
    Returns list of successfully downloaded PDF paths.
    """
    if tickers is None:
        from config import COMPANIES
        tickers = list(COMPANIES.keys())
    if years is None:
        years = DEFAULT_YEARS

    downloaded: list[Path] = []

    for ticker in tickers:
        print(f"\n[{ticker}] Scraping CafeF...")
        links = _get_pdf_links_playwright(ticker, years)

        if not links:
            print(f"  ✗ No matching PDFs found for {ticker}")
            continue

        for item in sorted(links, key=lambda x: x["year"]):
            dest = DATA_RAW / f"{item['ticker']}_{item['year']}.pdf"
            success = _download_pdf(item["url"], dest)
            if success:
                downloaded.append(dest)
            time.sleep(0.5)

    print(f"\n── Download complete: {len(downloaded)} PDFs ──")
    return downloaded


if __name__ == "__main__":
    run()
