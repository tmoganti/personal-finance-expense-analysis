# Personal Finance & Expense Analysis — OSEMN Portfolio Project

A portfolio-ready personal finance analysis built on the **OSEMN** framework
(Obtain → Scrub → Explore → Model → iNterpret) over my own transaction feed.

It answers four questions:

1. How much am I spending each month?
2. Which categories take most of my money?
3. Are my expenses increasing over time?
4. Where can I reduce spending?

And, because single-month spikes can mislead, it separates **recurring baseline
spending** from **one-off charges** (government fees, annual renewals, tourist
days, bimonthly utility lumps) so the answers reflect habits, not outliers.

## Dataset

Real personal transactions covering **Dec 12 2025 → Mar 24 2026**
(≈3.5 months, **121 transactions**, **$4,506.53 total**).

| Source                   | Rows | Span                       |
|--------------------------|-----:|----------------------------|
| `expenses.csv`           | 121  | 2025-12-12 → 2026-03-24    |
| `extracted_from_pdfs.csv`| 121  | 2025-12-12 → 2026-03-24    |

Columns in `expenses.csv`:
`TransactionID, Date, Year, Month, YearMonth, Description, Category,
Category_Raw, Amount, Payment Method, Weekday, IsOneOff`.

- `Category_Raw` preserves the original label before consolidation so the
  remap is auditable (e.g. `Supermarkets` → `Groceries`,
  `Gas/Convenience` → `Gasoline`, `Online Services` → `Subscriptions`).
- `IsOneOff` (`Yes`/`No`) flags non-recurring transactions: government
  fees, a one-time tourist visit (Summit One Vanderbilt NY), the
  bimonthly Eversource utility lump, and the annual Mint Mobile
  renewal. Flagged one-offs total **$1,797.61 (39.9% of spend)**, which
  is why the recurring baseline is a much better read on monthly habits.

## Deliverables

| File | What it is | How to use |
|---|---|---|
| `Personal_Finance_Dashboard.xlsx` | Interactive Excel dashboard — 3 filter dropdowns, 5 KPIs, 5 charts incl. a 2-series Month-over-Month (Total vs Recurring), Key Insights panel. | Open in Microsoft Excel. Use the **Month / Category / Payment Method** dropdowns at the top — every KPI and chart updates automatically. |
| `Personal_Finance_Dashboard_Numbers.xlsx` | Same dashboard rebuilt with A1-style ranges and no `_xlfn.` prefixes so it imports into **Apple Numbers** without `#NAME?` errors. | Open in Numbers (File → Open). Filter dropdowns become pop-up menus. |
| `Expense_Analysis_OSEMN.ipynb` | Fully-executed Jupyter notebook walking through OSEMN with pandas + matplotlib + seaborn. Outputs are pre-rendered; no kernel needed to read. | Open in Jupyter / VS Code to read; "Run All" to reproduce. |
| `Dashboard_Preview.png` | One-shot 1920×1080 composite of the full dashboard — KPI band + 5 charts + Insights footer. | LinkedIn post, slide thumbnail, portfolio README hero image. |
| `expenses.csv` | The cleaned, analysis-ready transaction feed (output of `scripts/clean_data.py`). | Swap in your own data — keep the same column names and the dashboard just works. |
| `extracted_from_pdfs.csv` | Raw output of the PDF → CSV extractor (no consolidation, no one-off flagging). | Audit trail that ties each txn back to its source statement PDF. |
| `figures/` | Standalone PNG exports of each notebook chart. | Slide decks, LinkedIn, etc. |
| `scripts/` | The six scripts that produced everything above (see below). | Re-run to rebuild any deliverable from scratch. |

## Dashboard layout

```
┌────────────────────────────────────────────────────────────────┐
│   PERSONAL FINANCE DASHBOARD · Dec 2025 – Mar 2026             │
├──────────┬──────────┬──────────────────────────────────────────┤
│ MONTH    │ CATEGORY │ PAYMENT METHOD   (dropdown filters)      │
├──────────┴──────────┴──────────────────────────────────────────┤
│ TOTAL │ MONTHLY │ TRANS.  │ TOP        │ AVG / TXN  (or        │
│ SPEND │ AVG     │         │ CATEGORY   │ LARGEST TXN in Excel) │
├────────────────────────────┬───────────────────────────────────┤
│  Spend by Category         │  Daily Spend Trend                │
├────────────┬───────────────┴───────────┬───────────────────────┤
│ Month-over │  Payment Method Share     │  Top 10 Merchants     │
│ -Month     │  (Total vs Recurring)     │                       │
├────────────┴───────────────────────────┴───────────────────────┤
│  KEY INSIGHTS                                                  │
└────────────────────────────────────────────────────────────────┘
```

## How the interactivity works

- The **Transactions** sheet holds every row as an Excel Table named `Transactions`.
- The **Summary_Tables** sheet holds small lookup tables (one per chart)
  driven by `SUMIFS` / `COUNTIFS` formulas that reference three filter
  cells on the Dashboard:
  - `Dashboard!B5` — Month filter (YearMonth e.g. `2026-01`)
  - `Dashboard!E5` — Category filter
  - `Dashboard!H5` — Payment Method filter
  - When a filter = `"All"`, the formula passes `"<>"` as the criterion
    (matches any non-blank) so the filter effectively disables itself.
- Charts point at the summary ranges, so flipping any dropdown instantly
  re-draws every chart and re-computes every KPI. Zero manual refresh.
- The **Month-over-Month** chart is a 2-series clustered column —
  **Total** (navy) next to **Recurring only** (orange) — so one-off
  charges stop drowning out the baseline. This uses a separate
  `SUMIFS(..., Transactions[IsOneOff], "No")` formula.

## Excel vs Numbers

Two variants ship to cover both apps:

- **Excel (`Personal_Finance_Dashboard.xlsx`)** uses structured table
  references (`Transactions[Amount]`) and `_xlfn.MAXIFS` for the
  *Largest Txn* KPI. Native in Excel 2013+.
- **Numbers (`Personal_Finance_Dashboard_Numbers.xlsx`)** uses A1
  ranges (`Transactions!$I$2:$I$122`) and swaps `MAXIFS` for
  `=IFERROR(TotalSpend / TransactionCount, 0)` (`Avg / Txn`) because
  Numbers doesn't resolve the `_xlfn.` prefix and shows `#NAME?` for
  `MAXIFS`. An "Open in Numbers" guide sheet documents what to click.

## OSEMN steps covered in the notebook

1. **Obtain** — load `expenses.csv`, describe the feed.
2. **Scrub** — null/duplicate/type checks, confirm the Dec-2026→Dec-2025
   typo fix from `clean_data.py`, audit the category consolidation
   (show every `Category_Raw` → `Category` remap), audit the `IsOneOff`
   flags (what's excluded and why).
3. **Explore** — headline KPIs, spend-by-category, daily trend, weekday
   pattern, payment-method breakdown, top merchants.
4. **Model** — month-over-month *Total vs Recurring* comparison,
   category × month pivot, 4-point linear trend (directional caveat),
   Needs / Wants bucket mapping adapted to the consolidated categories.
5. **iNterpret** — written findings & recommendations tied to the
   numbers (Restaurants is the biggest controllable lever, January's
   $2.17k peak drops to $1.05k once one-offs are stripped, BofA Credit
   carries ~76% of spend, etc.).

## Scripts

Everything in `scripts/` is idempotent — wipe the outputs and re-run to
rebuild from scratch.

| Script | What it does | Output |
|---|---|---|
| `scripts/extract_pdfs.py` | Reads every bank/credit-card PDF in this folder, detects the format (Discover · BofA Visa · BofA Checking), and extracts purchase rows into a unified CSV. Discover categories are kept as-issued; BofA credit-card rows get `Uncategorized` for manual mapping. BofA checking statements are intentionally left as a stub — the deposit/debit layout is different and belongs in a separate parser. | `extracted_from_pdfs.csv` |
| `scripts/clean_data.py` | Scrubs the raw CSV: trims headers, fixes 4 typo rows dated `2026-12-31` → `2025-12-31`, consolidates overlapping categories, flags one-off transactions, derives time features, assigns `TransactionID`. | `expenses.csv` |
| `scripts/build_dashboard.py` | Builds the Excel-native dashboard via `openpyxl`. Structured table refs + `_xlfn.MAXIFS`. | `Personal_Finance_Dashboard.xlsx` |
| `scripts/build_dashboard_numbers.py` | Builds the Numbers-compatible dashboard. A1 refs, no `_xlfn.`, Avg/Txn KPI, PieChart instead of DoughnutChart, pop-up menu data validation. | `Personal_Finance_Dashboard_Numbers.xlsx` |
| `scripts/build_notebook.py` | Hand-rolls the Jupyter `.ipynb` with all 27 cells pre-executed and all figures embedded as base64 PNGs. | `Expense_Analysis_OSEMN.ipynb` + `figures/*.png` |
| `scripts/render_preview.py` | Renders the single-image composite PNG preview of the dashboard for LinkedIn / slide decks. | `Dashboard_Preview.png` |

## Using your own data

Drop new statement PDFs next to the existing ones and re-run:

```bash
python scripts/extract_pdfs.py       # PDFs → extracted_from_pdfs.csv
python scripts/clean_data.py         # raw CSV → expenses.csv
python scripts/build_dashboard.py    # expenses.csv → Excel dashboard
python scripts/build_dashboard_numbers.py
python scripts/build_notebook.py
python scripts/render_preview.py
```

The column schema is the contract — as long as your feed has
`Date, Description, Category, Amount, Payment Method`, everything
downstream just works.

## Upstream sources

The **9 PDF statements** in this folder are the original source of truth:

- `January 2026.pdf`, `February 2026.pdf`, `March 2026.pdf` — Discover
  credit card monthly statements.
- `eStmt_2026-01-27.pdf`, `-02-27.pdf`, `-03-27.pdf` — BofA Visa
  Signature credit card statements.
- `eStmt_2026-01-23.pdf`, `-02-20.pdf`, `-03-24.pdf` — BofA SafeBalance
  checking account statements (not yet parsed; the deposit/debit layout
  needs a dedicated parser in a future iteration).

`scripts/extract_pdfs.py` ingests the first six and reconstructs
`expenses.csv` exactly (same 121 transactions, same $4,506.53 total) —
the end-to-end pipeline is fully reproducible from the raw PDFs.
