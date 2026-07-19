# EventLens — a no-code Excel event-study model

EventLens measures how abnormal 15 major market shocks (2001–2024) were across 8 US benchmark
series, using a formal event-study framework built **entirely in Excel formulas** — no VBA, no
add-ins, no code required to use it. Open the workbook, and 118,000+ live formulas do the rest.

**Grid:** 15 events × 8 benchmarks = 120 event studies (81 computed in the current release;
the rest auto-complete when benchmark data is extended — see *Extending the data* below).

## Headline results (CAR over the −5..+5 trading-day window)

| | |
|---|---|
| Deepest negative CAR | **WTI crude −62.3%*** on the COVID-19 declaration (12-03-2020) |
| Largest positive CAR | **VIX +84.3%*** on the COVID-19 declaration |
| Only event significant on every equity benchmark | US credit-rating downgrade (08-08-2011) |
| Small-cap divergence | Trump 2016 election: S&P 600 proxy **+10.6%***, large-caps under +4% |
| Demand vs supply shocks in oil | 9/11 −25.6%***, SVB −14.1%* (demand) vs Ukraine **+13.3%*** (supply) |

`* / ** / ***` = significant at 10% / 5% / 1%.

## Methodology

- **Model:** constant-mean-return (Brown & Warner 1985). The assets under study are broad
  benchmarks, so a Market-Model regression against the S&P 500 would absorb the event itself.
- **Returns:** daily log returns, additive across the window, so CAR is a clean sum.
- **Windows:** 120-trading-day estimation window, 10-day buffer, (−5, +5) event window.
  Every parameter lives once on the `Inputs` sheet as a named range.
- **Significance:** manual Brown–Warner t-statistics (`t = CAR / (σ·√11)`), two-tailed p-values.
  Excel's `T.TEST` is deliberately not used — it answers a two-sample question; event studies
  ask a one-sample one.
- **Compatibility:** INDEX/MATCH only. No XLOOKUP, FILTER, OFFSET, or other volatile/modern
  functions. Works in Excel 2010+ and LibreOffice.

## Workbook tour

`README` → `Inputs` → five `RawData_*` sheets → `Calc_Returns` → `Calc_EventStudies`
(120-row engine) → `Calc_Significance` (t-stats, p-values, source checks, coverage panel) →
`Dashboard` (gapless results table + headlines) → `MatrixView` (15×8 coverage matrices + chart).

Color convention: **blue** = hardcoded input · **black** = same-sheet formula ·
**green** = cross-sheet formula · **red** = flags/checks · **yellow** = paste areas.

## Data sources & attribution

| Series | Source |
|---|---|
| S&P 500 | GitHub mirrors of Yahoo `^GSPC` ([vijinho/sp500](https://github.com/vijinho/sp500) + second mirror) spliced with the S&P 500 index bundled in [skfolio](https://pypi.org/project/skfolio/); sources agree to 0.0000% on 5,537 overlapping days |
| Nasdaq Composite | GitHub mirror of Yahoo `^IXIC` (2011–2022) + ONEQ ETF scaled to index level (2005–2010) |
| Dow Jones | GitHub mirror of Yahoo `^DJI` (2011–2022) + DIA ETF scaled at the seam (2005–2010) |
| Russell 3000 / S&P 600 / CRSP Total Market | ETF proxies IWV / IJR / VTI, dividend-adjusted, from a GitHub mirror of the Stooq-lineage Kaggle "Huge Stock Market Dataset" (2005–2017) |
| WTI crude oil | [datasets/oil-prices](https://github.com/datasets/oil-prices) (EIA/FRED lineage, auto-updated) |
| VIX | [datasets/finance-vix](https://github.com/datasets/finance-vix) (CBOE, auto-updated) |

Raw per-series CSVs used to populate the workbook are in [`data/`](data/). Price data remains
subject to the original providers' terms; it is redistributed here, with attribution, for
educational purposes only.

## Integrity guardrails

- Every event-day return was validated against an independently curated event database before
  entering the model (e.g., Lehman S&P −4.71%, COVID S&P −9.51% — exact matches).
- The event database's narrative figures are **context/validation only** — they never feed a
  calculation. A `Source check` column recomputes day-0 returns from raw prices and flags any
  divergence beyond 0.25%.
- ETF-proxy segments are documented per date range; proxy-vs-index divergence on chaotic days
  (up to 1.4% for IJR on Lehman day) is flagged, not hidden.
- Studies that cannot be computed honestly are **excluded, never fabricated** — e.g.
  VTI × 9/11 (the fund did not exist yet).

## Extending the data

Each benchmark's `RawData_*` sheet accepts appended rows (Date, Close — oldest first).
Download links for every series are on the `Inputs` sheet. All 120 studies, the Dashboard
table, and the matrices recompute automatically.

## Disclaimer

Educational project. Nothing here is investment advice. No warranty of fitness for any
purpose; verify independently before relying on any figure.
