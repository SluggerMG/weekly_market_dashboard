# Weekly Market Pull — Claude Code instructions

When I ask you to "run the weekly market pull" (or refresh the market data /
grab this week's closes), follow the `weekly-market-pull` skill
(`.claude/skills/weekly-market-pull/SKILL.md`). It can also be invoked
directly with `/weekly-market-pull`.

## Notes
- Everything runs in this cloud environment — nothing installs on my computer.
- The script config (tickers, strip end year, Treasury toggle) is in the
  CONFIGURATION block near the top of weekly_market_report.py. If I ask to track
  something new, edit that block rather than hand-editing output.
