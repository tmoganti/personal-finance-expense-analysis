"""
Build Personal_Finance_Dashboard.xlsx from REAL expenses.csv.

Dataset spans ~4 months (Dec 2025 → Mar 2026, 121 rows, 20 categories,
2 payment methods) so the dashboard is adapted:

  • Month filter (YearMonth) replaces Year filter
  • 5 charts: Spend by Category, Daily Spend Trend, Payment Method,
    Top Merchants, Month-over-Month Comparison
  • KPI cards: Total Spend, Monthly Avg, Transactions, Top Category,
    Largest Transaction
"""
from pathlib import Path
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.chart import BarChart, LineChart, DoughnutChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.marker import DataPoint
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties

HERE = Path(__file__).resolve().parent
CSV  = HERE / "expenses.csv"
OUT  = HERE / "Personal_Finance_Dashboard.xlsx"

# Palette
NAVY, BLUE, SKY, ACCENT = "1F4E78", "2E75B6", "5B9BD5", "F79256"
LIGHT, BG, INK, GRAY, WHITE = "DDEBF7", "F2F7FC", "1F2937", "6B7280", "FFFFFF"
THIN = Side(style="thin", color="D1D5DB")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

def bold(sz=11, color=INK): return Font(name="Calibri", size=sz, bold=True, color=color)
def reg(sz=11,  color=INK): return Font(name="Calibri", size=sz, color=color)
def fill(h): return PatternFill("solid", start_color=h, end_color=h)
def center(wrap=False): return Alignment(horizontal="center", vertical="center", wrap_text=wrap)
def left(wrap=False):   return Alignment(horizontal="left",   vertical="center", wrap_text=wrap)

# ======================================================================
# 1. Load data & Transactions sheet
# ======================================================================
df = pd.read_csv(CSV, parse_dates=["Date"])
wb = Workbook()

tx = wb.active
tx.title = "Transactions"
cols = ["TransactionID", "Date", "Year", "Month", "YearMonth",
        "Description", "Category", "Category_Raw",
        "Amount", "Payment Method", "Weekday", "IsOneOff"]
tx.append(cols)
for r in df[cols].itertuples(index=False):
    tx.append(list(r))

last_row_tx = len(df) + 1
for row in range(2, last_row_tx + 1):
    tx.cell(row=row, column=2).number_format = "yyyy-mm-dd"
    tx.cell(row=row, column=9).number_format = '$#,##0.00'   # Amount moved col H→I

for i, w in enumerate([14, 12, 8, 8, 11, 32, 18, 15, 12, 16, 12, 10], start=1):
    tx.column_dimensions[get_column_letter(i)].width = w

for c in range(1, len(cols) + 1):
    cell = tx.cell(row=1, column=c)
    cell.font = Font(name="Calibri", size=11, bold=True, color=WHITE)
    cell.fill = fill(NAVY); cell.alignment = center(); cell.border = BORDER
tx.freeze_panes = "A2"

table = Table(displayName="Transactions",
              ref=f"A1:{get_column_letter(len(cols))}{last_row_tx}")
table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
tx.add_table(table)

# ======================================================================
# 2. Summary_Tables — formula-driven summaries
# ======================================================================
st = wb.create_sheet("Summary_Tables")
st.sheet_properties.tabColor = BLUE

# Filter cells on Dashboard:
#   B5 = YearMonth filter ("All" or "2026-01" etc.)
#   E5 = Category filter
#   H5 = Payment Method filter
F_YM  = 'IF(Dashboard!$B$5="All","<>",Dashboard!$B$5)'
F_CAT = 'IF(Dashboard!$E$5="All","<>",Dashboard!$E$5)'
F_PM  = 'IF(Dashboard!$H$5="All","<>",Dashboard!$H$5)'

def s_amt(extra_col="", extra_val=""):
    crit = [
        'Transactions[YearMonth]',      F_YM,
        'Transactions[Category]',       F_CAT,
        'Transactions[Payment Method]', F_PM,
    ]
    if extra_col:
        crit.extend([extra_col, extra_val])
    return '=SUMIFS(Transactions[Amount],' + ','.join(crit) + ')'

def c_amt(extra_col="", extra_val=""):
    crit = [
        'Transactions[YearMonth]',      F_YM,
        'Transactions[Category]',       F_CAT,
        'Transactions[Payment Method]', F_PM,
    ]
    if extra_col:
        crit.extend([extra_col, extra_val])
    return '=COUNTIFS(Transactions[Date],"<>",' + ','.join(crit) + ')'

# A-C: Category summary (all categories sorted by overall total for a sensible order)
cat_order = df.groupby("Category")["Amount"].sum().sort_values(ascending=True).index.tolist()
st["A1"], st["B1"], st["C1"] = "Category", "Total", "Count"
for i, cat in enumerate(cat_order, start=2):
    st.cell(row=i, column=1, value=cat)
    st.cell(row=i, column=2, value=s_amt('Transactions[Category]', f'"{cat}"'))
    st.cell(row=i, column=3, value=c_amt('Transactions[Category]', f'"{cat}"'))
cat_last = 1 + len(cat_order)

# E-G: YearMonth summary — Total and Recurring-only (strips IsOneOff="Yes")
yms = sorted(df["YearMonth"].unique().tolist())
st["E1"], st["F1"], st["G1"] = "YearMonth", "Total", "Recurring"
for i, ym in enumerate(yms, start=2):
    st.cell(row=i, column=5, value=ym)
    st.cell(row=i, column=6,
            value=f'=SUMIFS(Transactions[Amount],'
                  f'Transactions[YearMonth],E{i},'
                  f'Transactions[Category],{F_CAT},'
                  f'Transactions[Payment Method],{F_PM})')
    st.cell(row=i, column=7,
            value=f'=SUMIFS(Transactions[Amount],'
                  f'Transactions[YearMonth],E{i},'
                  f'Transactions[Category],{F_CAT},'
                  f'Transactions[Payment Method],{F_PM},'
                  f'Transactions[IsOneOff],"No")')
ym_last = 1 + len(yms)

# I-J: Payment Method summary
pms = sorted(df["Payment Method"].unique().tolist())
st["I1"], st["J1"] = "Payment Method", "Total"
for i, pm in enumerate(pms, start=2):
    st.cell(row=i, column=9,  value=pm)
    st.cell(row=i, column=10,
            value=f'=SUMIFS(Transactions[Amount],'
                  f'Transactions[Payment Method],I{i},'
                  f'Transactions[YearMonth],{F_YM},'
                  f'Transactions[Category],{F_CAT})')
pm_last = 1 + len(pms)

# M-N: Top 10 merchants (by overall spend; values filter-aware)
top_merch = df.groupby("Description")["Amount"].sum().sort_values(ascending=True).tail(10).index.tolist()
st["M1"], st["N1"] = "Merchant", "Total"
for i, m in enumerate(top_merch, start=2):
    st.cell(row=i, column=13, value=m)
    safe = m.replace('"', '""')
    st.cell(row=i, column=14, value=s_amt('Transactions[Description]', f'"{safe}"'))
merch_last = 1 + len(top_merch)

# Q-R: Daily spend (last 35 days of dataset → shows recent daily pattern)
daily = df.groupby("Date")["Amount"].sum().reset_index().sort_values("Date")
tail = daily.tail(40)
st["Q1"], st["R1"] = "Date", "Spend"
for i, (_, row) in enumerate(tail.iterrows(), start=2):
    st.cell(row=i, column=17, value=row["Date"].strftime("%m-%d"))
    st.cell(row=i, column=18,
            value=f'=SUMIFS(Transactions[Amount],'
                  f'Transactions[Date],"{row["Date"].strftime("%Y-%m-%d")}",'
                  f'Transactions[Category],{F_CAT},'
                  f'Transactions[Payment Method],{F_PM})')
daily_last = 1 + len(tail)

# Headers + formats
for row in st.iter_rows(min_row=1,
                        max_row=max(cat_last, ym_last, pm_last, merch_last, daily_last),
                        min_col=1, max_col=18):
    for cell in row:
        cell.font = reg(10)
        if cell.row == 1:
            cell.font = bold(10, WHITE); cell.fill = fill(NAVY); cell.alignment = center()
        if isinstance(cell.value, str) and cell.value.startswith("="):
            if cell.column in (2, 6, 7, 10, 14, 18):
                cell.number_format = '$#,##0.00;($#,##0.00);-'
            elif cell.column == 3:
                cell.number_format = '#,##0'

for col_letter, w in [("A", 22), ("B", 14), ("C", 10), ("E", 12), ("F", 14), ("G", 14),
                      ("I", 18), ("J", 14), ("M", 34), ("N", 14),
                      ("Q", 10), ("R", 14)]:
    st.column_dimensions[col_letter].width = w

# ======================================================================
# 3. Dashboard sheet
# ======================================================================
dash = wb.create_sheet("Dashboard", 0)
dash.sheet_properties.tabColor = NAVY
dash.sheet_view.showGridLines = False

col_widths = {"A": 2,  "B": 22, "C": 22, "D": 2,  "E": 22, "F": 22, "G": 2,
              "H": 22, "I": 22, "J": 2,  "K": 22, "L": 22, "M": 2,  "N": 22, "O": 22, "P": 2}
for col, w in col_widths.items():
    dash.column_dimensions[col].width = w

# Title
title_range = f"Dec {df['Date'].min():%Y} – {df['Date'].max():%b %Y}"
dash.merge_cells("B1:O2")
t = dash["B1"]
t.value = f"PERSONAL FINANCE DASHBOARD  ·  {title_range}"
t.font = Font(name="Calibri", size=22, bold=True, color=WHITE)
t.fill = fill(NAVY); t.alignment = center()
dash.row_dimensions[1].height = 24
dash.row_dimensions[2].height = 22

dash.merge_cells("B3:O3")
dash["B3"].value = "Interactive view — pick filters below to slice the data"
dash["B3"].font = Font(name="Calibri", size=11, italic=True, color=GRAY)
dash["B3"].fill = fill(BG); dash["B3"].alignment = center()
dash.row_dimensions[3].height = 18

# --- Filter row ---
def styled_filter(cell, label="All"):
    c = dash[cell]
    c.value = label
    c.font = bold(14, NAVY); c.fill = fill(LIGHT)
    c.alignment = center(); c.border = BORDER

dash.merge_cells("B5:C5"); styled_filter("B5", "All")
dash.merge_cells("E5:F5"); styled_filter("E5", "All")
dash.merge_cells("H5:I5"); styled_filter("H5", "All")
dash.merge_cells("B4:C4"); dash["B4"].value = "MONTH"
dash.merge_cells("E4:F4"); dash["E4"].value = "CATEGORY"
dash.merge_cells("H4:I4"); dash["H4"].value = "PAYMENT METHOD"
for ref in ("B4", "E4", "H4"):
    dash[ref].font = bold(10, GRAY); dash[ref].alignment = left(); dash[ref].fill = fill(WHITE)

def dv_list(values):
    csv_list = ",".join(values)
    dv = DataValidation(type="list", formula1=f'"{csv_list}"', allow_blank=True)
    dv.error = "Pick a value from the list"
    dv.prompt = "Click the arrow to filter"
    return dv

dv_ym  = dv_list(["All"] + yms)
dv_cat = dv_list(["All"] + cat_order[::-1])  # present in descending-total order
dv_pm  = dv_list(["All"] + pms)
dash.add_data_validation(dv_ym);  dv_ym.add("B5")
dash.add_data_validation(dv_cat); dv_cat.add("E5")
dash.add_data_validation(dv_pm);  dv_pm.add("H5")

dash.row_dimensions[4].height = 18
dash.row_dimensions[5].height = 26
dash.row_dimensions[6].height = 10

# --- KPI row ---
# Number of selected months (for Monthly Avg). Uses Summary_Tables F column (count months with spend>0)
n_months_formula = f'COUNTIF(Summary_Tables!$F$2:$F${ym_last},">0")'

sumifs_all = (
    '=SUMIFS(Transactions[Amount],'
    f'Transactions[YearMonth],{F_YM},'
    f'Transactions[Category],{F_CAT},'
    f'Transactions[Payment Method],{F_PM})'
)
countifs_all = (
    '=COUNTIFS(Transactions[Date],"<>",'
    f'Transactions[YearMonth],{F_YM},'
    f'Transactions[Category],{F_CAT},'
    f'Transactions[Payment Method],{F_PM})'
)

# Largest transaction formula: max Amount where criteria met.
# Use _xlfn.MAXIFS prefix — MAXIFS was added post-Excel 2007 and openpyxl
# requires the _xlfn. namespace for newer functions or LibreOffice returns #NAME?.
maxifs_all = (
    '=_xlfn.MAXIFS(Transactions[Amount],'
    f'Transactions[YearMonth],{F_YM},'
    f'Transactions[Category],{F_CAT},'
    f'Transactions[Payment Method],{F_PM})'
)

kpis = [
    ("B", "C", "TOTAL SPEND",     sumifs_all,                                                     '$#,##0.00'),
    ("E", "F", "MONTHLY AVG",     f'=IFERROR(B8/{n_months_formula},0)',                           '$#,##0.00'),
    ("H", "I", "TRANSACTIONS",    countifs_all,                                                   '#,##0'),
    ("K", "L", "TOP CATEGORY",
        f'=INDEX(Summary_Tables!$A$2:$A${cat_last},MATCH(MAX(Summary_Tables!$B$2:$B${cat_last}),Summary_Tables!$B$2:$B${cat_last},0))', '@'),
    ("N", "O", "LARGEST TXN",     maxifs_all,                                                     '$#,##0.00'),
]

for col1, col2, label, formula, fmt in kpis:
    dash.merge_cells(f"{col1}7:{col2}7")
    lcell = dash[f"{col1}7"]
    lcell.value = label; lcell.font = bold(10, WHITE); lcell.fill = fill(NAVY); lcell.alignment = center()
    dash.merge_cells(f"{col1}8:{col2}10")
    vcell = dash[f"{col1}8"]
    vcell.value = formula
    vcell.font = Font(name="Calibri", size=20, bold=True, color=NAVY)
    vcell.fill = fill(LIGHT); vcell.alignment = center()
    vcell.number_format = fmt; vcell.border = BORDER

for r in (7, 8, 9, 10):
    dash.row_dimensions[r].height = 18 if r == 7 else 20
dash.row_dimensions[11].height = 8

# --- Chart section layout ---
def section_title(ref, text):
    dash[ref] = text
    dash[ref].font = bold(12, NAVY); dash[ref].fill = fill(WHITE); dash[ref].alignment = left()

section_title("B12", "SPEND BY CATEGORY")
section_title("H12", "DAILY SPEND TREND")
section_title("B25", "MONTH-OVER-MONTH")
section_title("H25", "PAYMENT METHOD")
section_title("L25", "TOP MERCHANTS")

for r in range(12, 38):
    dash.row_dimensions[r].height = 16

# Spend by Category (horizontal bar)
bar1 = BarChart(); bar1.type = "bar"; bar1.style = 11
bar1.title = "Spend by Category"
bar1.x_axis.number_format = '"$"#,##0'
bar1.y_axis.majorGridlines = None
bar1.legend = None; bar1.height = 7.5; bar1.width = 13
bar1.add_data(Reference(st, min_col=2, min_row=1, max_row=cat_last, max_col=2), titles_from_data=True)
bar1.set_categories(Reference(st, min_col=1, min_row=2, max_row=cat_last))
for s in bar1.series:
    s.graphicalProperties = GraphicalProperties(solidFill=BLUE)
dash.add_chart(bar1, "B13")

# Daily spend (last 40 days)
line1 = LineChart(); line1.style = 12
line1.title = "Daily Spend (recent days)"
line1.y_axis.number_format = '"$"#,##0'
line1.y_axis.majorGridlines = None
line1.height = 7.5; line1.width = 19
line1.add_data(Reference(st, min_col=18, min_row=1, max_row=daily_last, max_col=18), titles_from_data=True)
line1.set_categories(Reference(st, min_col=17, min_row=2, max_row=daily_last))
line1.legend = None
for s in line1.series:
    s.graphicalProperties = GraphicalProperties(solidFill=BLUE)
    s.graphicalProperties.line = LineProperties(solidFill=BLUE, w=22000)
    s.smooth = False
dash.add_chart(line1, "H13")

# Month-over-month (column) — 2 series: Total vs Recurring-only
bar3 = BarChart(); bar3.type = "col"; bar3.style = 11
bar3.title = "Monthly Spend — Total vs Recurring"
bar3.y_axis.number_format = '"$"#,##0'
bar3.y_axis.majorGridlines = None
bar3.height = 7.5; bar3.width = 13
bar3.add_data(Reference(st, min_col=6, min_row=1, max_row=ym_last, max_col=7),
              titles_from_data=True)
bar3.set_categories(Reference(st, min_col=5, min_row=2, max_row=ym_last))
bar3.grouping = "clustered"
# legend on top for readability
from openpyxl.chart.legend import Legend
bar3.legend = Legend()
bar3.legend.position = "t"
colors = [NAVY, ACCENT]
for s, c in zip(bar3.series, colors):
    s.graphicalProperties = GraphicalProperties(solidFill=c)
dash.add_chart(bar3, "B26")

# Payment Method donut
donut = DoughnutChart(); donut.style = 10
donut.title = "Payment Method Share"
donut.height = 7.5; donut.width = 10
donut.add_data(Reference(st, min_col=10, min_row=1, max_row=pm_last, max_col=10), titles_from_data=True)
donut.set_categories(Reference(st, min_col=9, min_row=2, max_row=pm_last))
slice_colors = [BLUE, ACCENT, SKY, GRAY]
pts = []
for i, c in enumerate(slice_colors[:pm_last - 1]):
    p = DataPoint(idx=i)
    p.graphicalProperties = GraphicalProperties(solidFill=c)
    pts.append(p)
donut.series[0].data_points = pts
donut.series[0].dLbls = DataLabelList(showPercent=True, showSerName=False, showCatName=False,
                                      showVal=False, showLegendKey=False)
dash.add_chart(donut, "H26")

# Top merchants (bar)
bar2 = BarChart(); bar2.type = "bar"; bar2.style = 11
bar2.title = "Top 10 Merchants"
bar2.x_axis.number_format = '"$"#,##0'
bar2.y_axis.majorGridlines = None
bar2.height = 7.5; bar2.width = 11
bar2.add_data(Reference(st, min_col=14, min_row=1, max_row=merch_last, max_col=14), titles_from_data=True)
bar2.set_categories(Reference(st, min_col=13, min_row=2, max_row=merch_last))
bar2.legend = None
for s in bar2.series:
    s.graphicalProperties = GraphicalProperties(solidFill=ACCENT)
dash.add_chart(bar2, "L26")

# --- Insights panel ---
total = df["Amount"].sum()
n_months = df["YearMonth"].nunique()
cat_sum = df.groupby("Category")["Amount"].sum().sort_values(ascending=False)
pm_sum  = df.groupby("Payment Method")["Amount"].sum()
top_merch_name = df.groupby("Description")["Amount"].sum().idxmax()
top_merch_amt  = df.groupby("Description")["Amount"].sum().max()
bofa_share     = pm_sum.get("BofA Credit", 0) / total * 100
restaurants_coffee = cat_sum.get("Restaurants", 0) + cat_sum.get("Coffee", 0)

one_off_total = df.loc[df["IsOneOff"]=="Yes","Amount"].sum()
recur_total   = total - one_off_total
one_off_pct   = one_off_total / total * 100

# Top CATEGORY among recurring-only spend (more honest headline)
recur_cat_sum = df.loc[df["IsOneOff"]=="No"].groupby("Category")["Amount"].sum().sort_values(ascending=False)
top_recur_cat = recur_cat_sum.index[0]
top_recur_cat_val = recur_cat_sum.iloc[0]

insights = [
    f"• Total outflow across {n_months} months: ${total:,.2f} — one-off charges account for ${one_off_total:,.0f} ({one_off_pct:.0f}%), leaving ${recur_total:,.0f} of recurring spend.",
    f"• Recurring top category: {top_recur_cat} at ${top_recur_cat_val:,.0f} — this is the realistic lever for saving.",
    f"• Restaurants + Coffee combined: ${restaurants_coffee:,.0f} — a sizeable, fully-recurring discretionary slice.",
    f"• BofA Credit carries {bofa_share:.0f}% of your spend; Discover Card carries {100-bofa_share:.0f}%.",
    f"• Biggest single merchant is '{top_merch_name}' at ${top_merch_amt:,.2f}.",
    "• Month-over-Month chart: orange bar strips USCIS fee, Mint Mobile annual, tourist visits, and Eversource bill to reveal the steady baseline.",
]

INSIGHT_START = 39
dash.merge_cells(f"B{INSIGHT_START}:O{INSIGHT_START}")
dash[f"B{INSIGHT_START}"].value = "KEY INSIGHTS"
dash[f"B{INSIGHT_START}"].font = bold(12, WHITE)
dash[f"B{INSIGHT_START}"].fill = fill(NAVY); dash[f"B{INSIGHT_START}"].alignment = left()
dash.row_dimensions[INSIGHT_START].height = 22

for i, text in enumerate(insights, start=INSIGHT_START + 1):
    dash.merge_cells(f"B{i}:O{i}")
    c = dash[f"B{i}"]
    c.value = text; c.font = reg(11); c.alignment = left(wrap=True); c.fill = fill(BG)
    dash.row_dimensions[i].height = 22

# Background fill
for row in range(1, 47):
    for col in range(1, 18):
        cell = dash.cell(row=row, column=col)
        if cell.fill.fgColor.rgb in (None, "00000000", "FFFFFFFF"):
            cell.fill = fill(WHITE)

# Page setup
dash.page_setup.orientation = dash.ORIENTATION_LANDSCAPE
dash.page_setup.paperSize   = dash.PAPERSIZE_TABLOID
dash.page_setup.fitToWidth  = 1
dash.page_setup.fitToHeight = 1
dash.sheet_properties.pageSetUpPr.fitToPage = True
dash.print_options.horizontalCentered = True
dash.page_margins.left = dash.page_margins.right = 0.25
dash.page_margins.top  = dash.page_margins.bottom = 0.25
dash.print_area = "A1:P46"

# ======================================================================
# 4. Pivot_Ready sheet
# ======================================================================
pr = wb.create_sheet("Pivot_Ready")
pr.sheet_properties.tabColor = "A9D08E"
pr["A1"] = "Pivot-Ready View"
pr["A1"].font = bold(14, NAVY)
pr["A3"] = (
    "Build your own PivotTable here:\n"
    "1. Click any cell in the Transactions table (next sheet).\n"
    "2. Insert → PivotTable → place it here.\n"
    "3. Add Slicers from PivotTable Analyze → Insert Slicer."
)
pr["A3"].alignment = left(wrap=True); pr["A3"].font = reg(11, GRAY)
pr.merge_cells("A3:H8")
pr.column_dimensions["A"].width = 18
for col in "BCDEFGH":
    pr.column_dimensions[col].width = 15

wb.active = 0
wb.save(OUT)
print(f"Saved {OUT} ({OUT.stat().st_size/1024:.1f} KB)")
