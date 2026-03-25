# Saks Filed Claims Scraper

Scrapes all filed claims from the [Stretto Saks bankruptcy case](https://cases.stretto.com/Saks/filed-claims/) and exports to a formatted Excel file.

## What it collects (28 fields per claim)

- Claim No., Creditor Name, Creditor Address, Debtor Name, Date Filed, Claim Status, Schedule No.
- Current/Filed/Schedule amount breakdowns (General Unsecured, Priority, Secured, Admin Priority, Total)
- Proof of Claim PDF link
- Notice Parties
- Objection, Transfer, Withdrawal, and Stipulation History

## Usage

### Step 1: Scrape (browser console)

1. Open https://cases.stretto.com/Saks/filed-claims/ in Chrome
2. Press **F12** → **Console** tab
3. Paste the contents of `scrape_saks_claims.js` and press Enter
4. Wait for it to finish (~60-90 min, fully autonomous)
5. Two files auto-download: `saks_claims_data.json` (backup) and `Saks_Filed_Claims.xls`

The scraper handles WAF token expiry automatically via hidden iframe refresh. State is saved to `localStorage` — if anything goes wrong, just re-paste the script to resume.

### Step 2 (optional): Generate a proper .xlsx

The `.xls` from the browser is an HTML-based Excel file. For a native `.xlsx` with better formatting:

```bash
pip install openpyxl
python format_saks_claims.py saks_claims_data.json Saks_Filed_Claims.xlsx
```

## Technical Details

- Site uses AWS WAF + Google reCAPTCHA v3 — requires a real browser session
- Scraper throttles requests (2 parallel, 2s delays, 30s rest breaks every 50 claims) to avoid WAF blocks
- Signed API URLs expire ~20 min; scraper auto-refreshes via hidden iframe
- State persists in `localStorage` for crash recovery
