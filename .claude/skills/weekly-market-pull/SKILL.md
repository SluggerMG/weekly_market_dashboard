---
name: weekly-market-pull
description: Refresh this week's market data (indices, FX, equities, NG/WTI strip, Treasury curve) and write it into Weekly_Market_Report.xlsx. Use when the user asks to run the weekly market pull, refresh the market data, or grab this week's closes.
---

# Weekly Market Pull

1. Install dependencies (safe to run every time):
   ```
   pip install yfinance openpyxl requests
   ```

2. Run the script, writing the workbook into the repo so it can be committed:
   ```
   python weekly_market_report.py Weekly_Market_Report.xlsx
   ```

3. Read the script's console output and report:
   - the week-of date and the pull timestamp (Central Time)
   - how many natural gas and WTI contracts returned a quote
   - whether any section came back empty

4. **Connectivity check.** If the run shows prices coming through, great — the
   network allowlist is working. If instead you see errors mentioning
   `host_not_allowed`, `403 Forbidden`, connection resets, or every price is
   `N/A`, then the environment could NOT reach Yahoo Finance / Treasury. In
   that case, stop and say plainly that the allowed-domains configuration did
   not take effect — don't try to work around it. Remove any workbook/sheet
   produced from the failed run rather than committing it. That's the
   specific thing being tested.

5. If the workbook was produced successfully (real prices, not N/A), commit
   it so it can be downloaded:
   ```
   git add Weekly_Market_Report.xlsx
   git commit -m "Weekly market data <week-of date>"
   ```
   Then report the branch name so the file can be grabbed from GitHub.

## Notes
- Everything runs in this cloud environment — nothing installs on the user's
  computer.
- The script config (tickers, strip end year, Treasury toggle) is in the
  CONFIGURATION block near the top of weekly_market_report.py. If asked to
  track something new, edit that block rather than hand-editing output.
