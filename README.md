# Securities Data Reconciliation Tool (Excel VBA)

An Excel VBA tool that reconciles security reference data from two vendor sources and flags every discrepancy in one click.

Investment data teams receive the same securities from multiple data providers. Before that data is used in client reporting, it has to be checked: do both sources agree on each security's coupon, maturity, currency and price, and is anything missing or duplicated? This tool automates that check.

## What it does

Click **Run Reconciliation** and the macro:

1. Loads Vendor B into a lookup (ISIN → row) and flags **duplicate ISINs**
2. Matches every Vendor A security to Vendor B by **ISIN**
3. Compares each field (name, asset class, currency, coupon, maturity, price, rating) and flags:
   - **Field mismatches** (e.g. coupon or maturity differs between vendors)
   - **Missing values** (e.g. a bond with no maturity date in Vendor B)
   - **Price mismatches**, with a tolerance rule: gaps over 25% are flagged as **price outliers**
4. Finds securities **missing from either vendor**
5. Writes every issue to a `Recon_Results` sheet, with a summary count by issue type
6. Highlights problem cells in Vendor B: red = field issue, pink = duplicate, yellow = only in Vendor B

## Data

- **Government bonds:** all 70 conventional UK gilts in issue, from the UK Debt Management Office (DMO) *Gilts in Issue* report (30 October 2026): real ISINs, names, coupons and redemption dates.
- **Corporate bonds and equities:** fictional issuers with correctly formatted ISINs, added to test multiple asset classes.
- **Prices:** simulated (the DMO report contains reference data, not prices).

### Preprocessing of the DMO report
- Removed titles, summary lines, section labels (Ultra-Short / Short / Medium / Long) and footnotes
- Kept conventional gilts only; index-linked gilts were excluded because their values carry an inflation uplift and need separate handling
- Extracted the coupon from the gilt name (e.g. `4 1/8%` → 4.125, `3¾%` → 3.75), since the report has no separate coupon field
- Converted redemption dates to date values and standardised all securities to one schema

## Testing

Vendor B is a copy of Vendor A with **18 deliberately planted errors** across 9 issue types, documented in `answer_key.csv`. The macro detects all 18.

| Issue type | Planted | Detected |
|---|---|---|
| Price mismatch | 4 | 4 |
| Missing maturity in Vendor B | 3 | 3 |
| Coupon mismatch | 2 | 2 |
| Maturity mismatch | 2 | 2 |
| Missing from Vendor B | 2 | 2 |
| Missing from Vendor A | 2 | 2 |
| Currency mismatch | 1 | 1 |
| Price outlier (>25% gap) | 1 | 1 |
| Duplicate ISIN in Vendor B | 1 | 1 |
| **Total** | **18** | **18** |

## Techniques used

- `Scripting.Dictionary` for fast ISIN lookups instead of repeated row-by-row searches
- Dynamic last-row detection, so the tool works for any number of securities
- Nested loops (securities × fields) with type-safe comparison of text, numbers and dates
- Tolerance-based price checks to separate small mismatches from outliers
- Automatic results sheet, summary table and colour-coded highlighting

## How to use

1. Open `Securities_Recon_Tool.xlsm` and enable macros
2. Paste vendor data into the `Vendor_A` and `Vendor_B` sheets (same column layout)
3. Click **Run Reconciliation** on the Vendor_A sheet
4. Review issues on the `Recon_Results` sheet

## Screenshots

![Vendor B with highlighted discrepancies](vendor_b_highlights.png)
![Reconciliation results and summary](recon_results.png)

## Source

UK Debt Management Office, Gilts in Issue report: https://www.dmo.gov.uk/data/gilt-market/gilts-in-issue/
