**Subject:** Submission – Project 3: Multi-Company Financial Statement Normalizer

---

Hi [Name],

Please find below my submission for **Project 3 – Multi-Company Financial Statement Normalizer**.

---

**Overview**

I built an end-to-end Python pipeline that automatically downloads, extracts, and normalizes Vietnamese annual report PDFs from CafeF, then exports a structured Excel comparison workbook across multiple companies and years.

The three companies selected are all in the **Vietnamese steel sector**:

| Ticker | Company |
|--------|---------|
| HPG | Hoa Phat Group |
| HSG | Hoa Sen Group |
| NKG | Nam Kim Steel |

Years covered: **2022, 2023, 2024**

---

**What the tool produces**

- A per-company Excel workbook for each company with 9 sheets: formatted Income Statement, Balance Sheet, and Cash Flow; English-translated raw line items; and the original Vietnamese raw lines
- A **comparison workbook** (`comparison.xlsx`) with side-by-side financial ratios across all three companies and all three years, covering profitability, leverage, liquidity, efficiency, and cash flow quality
- An auto-generated **Insights sheet** in plain English flagging key observations and warning signals

---

**Deliverables**

- **GitHub repository:** https://github.com/kylenguyen291/20in20-financial-normalizer
- **Setup instructions:** See `README.md` in the repository
- **API key:** Provided in `api_key.txt` in the repository
- **Writeup:** `project_writeup.docx` in the repository (problem framing, architecture, tools, cost, edge cases, and next steps)

---

**Running the tool**

```bash
git clone https://github.com/kylenguyen291/20in20-financial-normalizer.git
cd 20in20-financial-normalizer
pip install -r requirements.txt
playwright install chromium
# Copy the API key from api_key.txt into your .env file as ANTHROPIC_API_KEY=...
python main.py
```

When prompted, enter the tickers (HPG, HSG, NKG) and years (2022, 2023, 2024). The full run takes approximately 10–15 minutes and costs ~$1.00–$1.50 in API credits.

---

**Cost summary**

- Development and testing: ~$20 total
- Per production run (3 companies × 3 years): ~$1.00–$1.50

---

Please let me know if you have any questions or need me to walk through the output.

Best regards,
Long
