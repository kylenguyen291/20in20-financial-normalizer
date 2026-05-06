"""
extractor.py
────────────
Extracts raw text from a Vietnamese financial statement PDF.

Strategy:
1. Try pdfplumber text/table extraction (works for digital PDFs)
2. If all pages return empty → PDF is a scanned image
   → Convert pages to PNG images via PyMuPDF
   → Return images for Claude vision processing in normalizer.py

Returns either:
  - str: extracted text (digital PDF)
  - list[dict]: list of base64 image dicts (scanned PDF)
  - None: total failure
"""

import base64
from pathlib import Path

import pdfplumber

from config import DATA_PROCESSED


def _extract_table_text(page) -> str:
    lines = []
    tables = page.extract_tables()
    for table in tables:
        for row in table:
            cleaned = [cell.strip() if cell else "" for cell in row]
            lines.append("\t".join(cleaned))
        lines.append("")
    return "\n".join(lines)


def _extract_page_text(page) -> str:
    text = page.extract_text()
    return text.strip() if text else ""


def _pdf_to_images(pdf_path: Path, max_pages: int = 25, dpi: int = 72) -> list[dict]:
    """
    Convert first `max_pages` of a PDF to base64-encoded PNG images.
    Returns list of Anthropic vision content dicts.
    """
    try:
        import fitz  # PyMuPDF
    except ImportError:
        print("  ✗ PyMuPDF not installed. Run: pip install PyMuPDF")
        return []

    doc = fitz.open(str(pdf_path))
    total = len(doc)
    # Financial tables are in the first ~30 pages of a standalone FS PDF
    pages_to_convert = min(max_pages, total)
    print(f"    Converting {pages_to_convert}/{total} pages to images (DPI={dpi})...")

    images = []
    mat = fitz.Matrix(dpi / 72, dpi / 72)

    for i in range(pages_to_convert):
        page = doc[i]
        pix = page.get_pixmap(matrix=mat)
        img_bytes = pix.tobytes("png")
        img_b64 = base64.standard_b64encode(img_bytes).decode()
        images.append({
            "type": "image",
            "source": {
                "type": "base64",
                "media_type": "image/png",
                "data": img_b64,
            },
        })

    doc.close()
    print(f"    ✓ {len(images)} page images ready for Claude vision")
    return images


def _is_drm_watermark(text: str, uniqueness_threshold: float = 0.10) -> bool:
    """
    Detect PDFs where every page carries the same DRM / eoffice watermark as
    the only extractable text.  pdfplumber sees real characters, passes the
    MIN_CHARS check, but every line is identical — Claude receives N copies of
    the same sentence and returns empty JSON.

    Heuristic: strip page-header lines ("--- Page N ---") then compare the
    number of *unique* content lines to total content lines.  If fewer than
    `uniqueness_threshold` (10%) of lines are unique, the content is repetitive
    junk and the PDF should be processed as a scanned image instead.
    """
    lines = [
        l.strip()
        for l in text.splitlines()
        if l.strip() and not l.startswith("--- Page")
    ]
    if not lines:
        return False
    ratio = len(set(lines)) / len(lines)
    return ratio < uniqueness_threshold


def extract(pdf_path: Path, force: bool = False):
    """
    Extract content from a PDF.

    Returns:
      - str:        extracted text (digital PDF) — cached to .txt
      - list[dict]: base64 image list (scanned PDF or mixed PDF) — not cached (too large)
      - None:       failure

    Four PDF types handled:
      1. Pure digital   — all pages have real varied text → text mode
      2. Pure scanned   — all pages empty                → vision mode
      3. Mixed          — some pages text, some scanned  → vision mode
      4. DRM/watermark  — every page has the same eoffice/DRM stamp as the
                          only extractable text; looks digital but content is
                          junk → vision mode
    """
    cache_path = DATA_PROCESSED / (pdf_path.stem + ".txt")

    # Return cached text if available — but still validate it isn't DRM junk
    if cache_path.exists() and not force:
        cached = cache_path.read_text(encoding="utf-8")
        if _is_drm_watermark(cached):
            print(f"  ⚠ Cached text is DRM/watermark junk — discarding cache, switching to vision")
            cache_path.unlink()          # delete the bad cache so it isn't reused
            return _pdf_to_images(pdf_path)
        print(f"  → Using cached extraction: {cache_path.name}")
        return cached

    print(f"  Extracting: {pdf_path.name}")

    all_text: list[str] = []
    page_char_counts: list[int] = []   # per-page char count for mixed-PDF detection

    # Characters per page below this = treat the page as "sparse / scanned"
    PAGE_MIN_CHARS = 50

    try:
        with pdfplumber.open(pdf_path) as pdf:
            total = len(pdf.pages)
            print(f"    {total} pages — trying text extraction...")

            for i, page in enumerate(pdf.pages, start=1):
                table_text = _extract_table_text(page)
                if table_text.strip():
                    all_text.append(f"--- Page {i} ---\n{table_text}")
                    page_char_counts.append(len(table_text.strip()))
                    continue
                raw = _extract_page_text(page)
                if raw:
                    all_text.append(f"--- Page {i} ---\n{raw}")
                    page_char_counts.append(len(raw.strip()))
                else:
                    page_char_counts.append(0)

    except Exception as e:
        print(f"  ✗ pdfplumber error: {e}")
        return None

    full_text = "\n\n".join(all_text)
    total_pages = len(page_char_counts)
    sparse_pages = sum(1 for c in page_char_counts if c < PAGE_MIN_CHARS)

    # ── Thresholds ────────────────────────────────────────────────────────────
    # Total-document threshold: below this the whole doc is considered non-digital
    MIN_CHARS = 1500
    # Mixed-PDF thresholds: if EITHER condition is true alongside enough total
    # text, the PDF has significant scanned pages that text mode would miss.
    MIXED_ABS_THRESHOLD   = 3     # > 3 sparse pages → suspicious
    MIXED_RATIO_THRESHOLD = 0.15  # > 15% of pages sparse → suspicious

    if not full_text.strip() or len(full_text.strip()) < MIN_CHARS:
        # ── Case 1/2: pure scanned, or mixed with almost no extractable text ──
        if full_text.strip():
            print(f"  ⚠ Text too short ({len(full_text.strip())} chars) — mixed/partial scan")
        else:
            print(f"  ⚠ All {sparse_pages} pages empty — scanned PDF detected")
        print(f"  → Switching to Claude vision mode")
        return _pdf_to_images(pdf_path)

    # We have enough total text — now check for the mixed-PDF trap:
    # enough text was extracted from *some* pages, but other pages are scanned
    # images that pdfplumber silently skipped. Those skipped pages may contain
    # the actual financial tables.
    is_mixed = (
        sparse_pages > MIXED_ABS_THRESHOLD
        or (total_pages > 0 and sparse_pages / total_pages > MIXED_RATIO_THRESHOLD)
    )

    if is_mixed:
        # ── Case 3: mixed PDF ─────────────────────────────────────────────────
        rich_pages = total_pages - sparse_pages
        ratio_pct  = int(sparse_pages / total_pages * 100)
        print(
            f"  ⚠ Mixed PDF detected: {rich_pages} text-rich pages, "
            f"{sparse_pages} empty/scanned pages ({ratio_pct}% sparse)"
        )
        print(f"  → Switching to full vision mode to capture all content")
        return _pdf_to_images(pdf_path)

    # ── Case 4: DRM / eoffice watermark PDF ──────────────────────────────────
    # Looks digital (passes MIN_CHARS) but every page has the same stamp text.
    if _is_drm_watermark(full_text):
        unique = len({l.strip() for l in full_text.splitlines()
                      if l.strip() and not l.startswith("--- Page")})
        print(f"  ⚠ DRM/watermark PDF detected ({unique} unique line(s) across all pages)")
        print(f"  → Switching to Claude vision mode")
        return _pdf_to_images(pdf_path)

    # ── Case 1: pure digital PDF ──────────────────────────────────────────────
    cache_path.write_text(full_text, encoding="utf-8")
    print(f"  ✓ Text extracted ({len(full_text):,} chars) → {cache_path.name}")
    return full_text


def run(pdf_paths: list[Path] | None = None, force: bool = False) -> dict:
    """
    Extract from all PDFs. Returns dict keyed by stem.
    Values are either str (text) or list[dict] (images).
    """
    from config import DATA_RAW

    if pdf_paths is None:
        pdf_paths = sorted(DATA_RAW.glob("*.pdf"))

    if not pdf_paths:
        print("No PDFs found in data/raw/. Run downloader first.")
        return {}

    results = {}
    for path in pdf_paths:
        print(f"\n[{path.stem}] Extracting...")
        content = extract(path, force=force)
        if content is not None:
            results[path.stem] = content

    text_count = sum(1 for v in results.values() if isinstance(v, str))
    img_count  = sum(1 for v in results.values() if isinstance(v, list))
    print(f"\n── Extraction: {text_count} text, {img_count} vision, "
          f"{len(pdf_paths) - len(results)} failed ──")
    return results


if __name__ == "__main__":
    run()
