# Portfolio ROI, Withdrawal & Tax Planning Workbook (Canada / Ontario)

`Portfolio_Planner_CA.xlsx` is the deliverable — a self-contained planning workbook
that imports cleanly into Google Sheets (File → Import → Upload, or upload to Drive
and open with Sheets). Everything except the yellow input cells is formula-driven,
so it keeps recalculating after import.

## Tabs

| Tab | Purpose |
| --- | --- |
| **Dashboard** | Year selector, KPI cards (balance, withdrawn, tax, dividends, net after-tax return, depletion indicator), ±1% sensitivity table, 4 charts |
| **Inputs** | The only tab to edit: portfolio value, horizon, 6 sub-portfolios with per-account splits, contributions, withdrawals, funding order, tax settings |
| **Projection** | Annual summary by account type + Withdrawals Summary block |
| **Tax Estimate** | Per-year CRA federal + Ontario tax (2026 rules): dividend gross-ups & credits, 50% capital-gains inclusion, RRSP as ordinary income, TFSA exempt; standalone and stacked views |
| **Engine (Monthly)** | One row per month; every figure on other tabs traces here |
| **Assumptions & Notes** | All tax parameters (named cells, with sources & as-of dates), engine conventions, exclusions |
| hidden | Engine (No Withdrawals) — counterfactual line; Engine (−1%) / (+1%) — sensitivity; ChartData |

## Scripts

- `build_workbook.py` — generates the workbook (`python3 build_workbook.py out.xlsx`). Requires `openpyxl`.
- `reference_model.py` — independent Python implementation of the engine + tax pipeline.
- `verify_workbook.py` — recalculates the workbook headlessly with LibreOffice Calc
  (`apt install libreoffice-calc`) and checks ~60 assertions: engine/tax/projection/dashboard
  values vs the reference model, the spec's acceptance test ($2,000/month from month 61 on a
  240-month horizon flows through every tab), and absence of error cells.
- `edge_tests.py` — same harness for edge cases: zero withdrawal, months-unit horizon,
  annual withdrawal frequency, DRIP off, routed contributions, reversed funding order.

To update tax figures in a future year, edit the constants at the top of
`build_workbook.py` (or the named cells on the Assumptions tab of the file itself)
— both `build_workbook.py` and `reference_model.py` must agree for the tests to pass.
