# Weekly Market Pull — Claude Code instructions

When I ask you to "run the weekly market pull" (or refresh the market data /
grab this week's closes), do the following:

1. Install dependencies (safe to run every time):
   ```
   pip install yfinance openpyxl requests
   ```

2. Run the script, writing the workbook into the repo so it can be committed:
   ```
   python weekly_market_report.py Weekly_Market_Report.xlsx
   ```

3. Read the script's console output and tell me:
   - the week-of date and the pull timestamp (Central Time)
   - how many natural gas and WTI contracts returned a quote
   - whether any section came back empty

4. **Connectivity check.** If the run shows prices coming through, great — the
   network allowlist is working. If instead you see errors mentioning
   `host_not_allowed`, `403 Forbidden`, or every price is `N/A`, then the
   environment could NOT reach Yahoo Finance / Treasury. In that case, stop and
   tell me plainly that the allowed-domains configuration did not take effect —
   don't try to work around it. That's the specific thing we're testing.

5. If the workbook was produced successfully, commit it so I can download it:
   ```
   git add Weekly_Market_Report.xlsx
   git commit -m "Weekly market data <week-of date>"
   ```
   Then tell me the branch name so I can grab the file from GitHub.

## Notes
- Everything runs in this cloud environment — nothing installs on my computer.
- The script config (tickers, strip end year, Treasury toggle) is in the
  CONFIGURATION block near the top of weekly_market_report.py. If I ask to track
  something new, edit that block rather than hand-editing output.
