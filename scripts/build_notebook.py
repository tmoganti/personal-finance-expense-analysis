"""
Hand-build a fully-executed .ipynb (v4 format) without jupyter installed.
Each code cell is executed in a shared namespace; stdout, DataFrame HTML
reprs, and matplotlib figures are captured into the notebook JSON so the
user sees outputs immediately on open.

This version walks the OSEMN stages over Tarun's REAL expense feed
(Dec 2025 → Mar 2026, 121 rows, extracted from Expenses.numbers).
"""
import base64, io, json, sys, traceback, contextlib, datetime
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

HERE = Path(__file__).resolve().parent
OUT_NB = HERE / "Expense_Analysis_OSEMN.ipynb"

# ---------- cell definitions ----------
C = []  # list of (kind, source)  kind ∈ {md, code}
md = lambda s: C.append(("md", s))
code = lambda s: C.append(("code", s.strip("\n")))

md("""# Personal Finance & Expense Analysis — OSEMN Walkthrough

**Author:** Tarun · **Framework:** OSEMN (Obtain → Scrub → Explore → Model → iNterpret)

**Goal.** Track and analyze monthly spending to answer:
- How much am I spending each month?
- Which categories take most of my money?
- Are my expenses increasing over time?
- Where can I reduce spending?

**Dataset.** Real personal transaction feed (`expenses.csv`) extracted from
`Expenses.numbers`, spanning **Dec 12 2025 → Mar 24 2026** (≈3.5 months,
121 rows). Fields: `TransactionID`, `Date`, `Description`, `Category`
(consolidated), `Category_Raw` (original label, kept for audit), `Amount`,
`Payment Method`, `IsOneOff` (flag for non-recurring charges).
""")

code("""
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from pathlib import Path

sns.set_theme(style="whitegrid", context="talk")
plt.rcParams["figure.figsize"] = (10, 5)
plt.rcParams["axes.titleweight"] = "bold"
pd.set_option("display.max_columns", 20)
pd.set_option("display.width", 160)
""")

md("## 1 · Obtain — load the transaction feed")

code("""
df = pd.read_csv("expenses.csv", parse_dates=["Date"])
print(f"Rows: {len(df):,}   Columns: {df.shape[1]}")
print(f"Date range: {df['Date'].min().date()} → {df['Date'].max().date()}")
df.head()
""")

md("## 2 · Scrub — clean, validate, enrich")

code("""
# Type / integrity checks
print("Missing values per column:")
print(df.isna().sum())

print(f"\\nDuplicate TransactionIDs: {df['TransactionID'].duplicated().sum()}")
print(f"Negative amounts (should be 0 for an expense feed): {(df['Amount'] < 0).sum()}")
print(f"Unique categories: {df['Category'].nunique()}")
print(f"Unique payment methods: {df['Payment Method'].nunique()}")
""")

code("""
# Fixing the one data issue caught during ingest:
#   4 rows in the source file were dated 2026-12-31 but sit BETWEEN
#   2025-12-30 and 2026-01-01 — obvious typo, corrected to 2025-12-31
#   in clean_data.py. Verify here:
print("Rows in Dec 2025:", (df['YearMonth'] == '2025-12').sum())
print("Rows in Dec 2026 (should be 0):", (df['YearMonth'] == '2026-12').sum())

# Time features are already present from clean_data.py, confirm:
df[["Date","Year","Month","YearMonth","Weekday"]].head()
""")

code("""
# Category consolidation audit — clean_data.py collapsed a handful of
# near-synonym labels (Supermarkets→Groceries, Gas/Convenience+Convenience
# Store→Gasoline, Merchandise→Shopping, Online Services→Subscriptions).
# The original label is preserved in Category_Raw so we can audit.
remap = df[df['Category'] != df['Category_Raw']]
(remap.groupby(['Category_Raw','Category']).size()
      .reset_index(name='rows')
      .sort_values('rows', ascending=False))
""")

code("""
# One-off flagging — clean_data.py marks transactions that are NOT part of the
# recurring monthly pattern (visa fees, court fines, annual subscription
# renewals, one-time tourist visits, bimonthly utility lumps). These rows
# distort the month-over-month picture when they land; flagging them lets us
# compute the "recurring baseline" separately.
one_off = df[df['IsOneOff'] == 'Yes']
print(f"One-off transactions: {len(one_off)} of {len(df)} "
      f"({len(one_off)/len(df)*100:.1f}% of rows)")
print(f"One-off amount:       ${one_off['Amount'].sum():,.2f} of "
      f"${df['Amount'].sum():,.2f} "
      f"({one_off['Amount'].sum()/df['Amount'].sum()*100:.1f}% of spend)")
(one_off.groupby(['Category','Description'])['Amount']
        .agg(['sum','count'])
        .round(2)
        .sort_values('sum', ascending=False))
""")

md("## 3 · Explore — where is the money going?")

code("""
# Headline KPIs
total       = df["Amount"].sum()
months      = df["YearMonth"].nunique()
monthly_avg = total / months
top_cat     = df.groupby("Category")["Amount"].sum().idxmax()
biggest_txn = df.loc[df["Amount"].idxmax()]

print(f"Total spend:                ${total:>12,.2f}")
print(f"Months covered:             {months:>12}")
print(f"Monthly average:            ${monthly_avg:>12,.2f}")
print(f"Transactions:               {len(df):>12,}")
print(f"Top category overall:       {top_cat}")
print(f"Largest single transaction: ${biggest_txn['Amount']:,.2f} "
      f"({biggest_txn['Description']}, {biggest_txn['Date'].date()})")
""")

code("""
# Spend by category
cat_totals = (df.groupby("Category")["Amount"]
                .agg(Total="sum", Count="count", AvgTxn="mean")
                .sort_values("Total", ascending=False)
                .round(2))
cat_totals["Share"] = (cat_totals["Total"] / cat_totals["Total"].sum()).map("{:.1%}".format)
cat_totals
""")

code("""
fig, ax = plt.subplots(figsize=(10, 7))
cat_totals["Total"].sort_values().plot.barh(ax=ax, color="#2E75B6")
ax.set_title("Total Spend by Category")
ax.set_xlabel("USD")
ax.set_ylabel("")
ax.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f"${x:,.0f}"))
for i, v in enumerate(cat_totals["Total"].sort_values()):
    ax.text(v, i, f"  ${v:,.0f}", va="center", fontsize=10)
plt.tight_layout()
plt.savefig("figures/01_spend_by_category.png", dpi=140, bbox_inches="tight")
plt.show()
""")

code("""
# Month-over-month trend (only 4 months of data so we skip rolling avg)
monthly = (df.groupby("YearMonth", as_index=False)["Amount"]
             .agg(Spend="sum", Transactions="count")
             .round(2))
monthly["AvgPerTxn"] = (monthly["Spend"] / monthly["Transactions"]).round(2)
monthly
""")

code("""
# Monthly spending — Total vs Recurring-only (strips one-offs).
# This is the single most important visual in the whole analysis: Jan looked
# like a $2.2k month until you strip the USCIS fee, Mint Mobile renewal, and
# Summit One Vanderbilt trip — the recurring baseline is ~$1k.
recur = df[df['IsOneOff'] == 'No']
monthly_recur = (recur.groupby('YearMonth')['Amount'].sum()
                      .reindex(monthly['YearMonth']).values)

fig, ax = plt.subplots(figsize=(11, 5.5))
x = np.arange(len(monthly))
w = 0.38
bars_total = ax.bar(x - w/2, monthly["Spend"],  w, label="Total",          color="#1F4E78")
bars_recur = ax.bar(x + w/2, monthly_recur,     w, label="Recurring only", color="#F79256")

ax.set_xticks(x); ax.set_xticklabels(monthly["YearMonth"])
ax.set_title("Monthly Spending — Total vs Recurring-only")
ax.set_ylabel("USD")
ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f"${x:,.0f}"))
ax.legend(loc="upper right")
for bars, vals in [(bars_total, monthly["Spend"]), (bars_recur, monthly_recur)]:
    for b, v in zip(bars, vals):
        ax.text(b.get_x() + b.get_width()/2, v, f"${v:,.0f}",
                ha="center", va="bottom", fontsize=9, fontweight="bold")
ax.set_ylim(0, max(monthly["Spend"].max(), max(monthly_recur)) * 1.18)
plt.tight_layout()
plt.savefig("figures/02_monthly_trend.png", dpi=140, bbox_inches="tight")
plt.show()
""")

code("""
# Daily spend over the whole window — shows the pattern of activity
daily = df.groupby("Date")["Amount"].sum()

fig, ax = plt.subplots(figsize=(12, 5))
ax.plot(daily.index, daily.values, marker="o", markersize=4, color="#2E75B6", linewidth=1.2)
ax.fill_between(daily.index, daily.values, alpha=0.25, color="#5B9BD5")
ax.set_title("Daily Spend")
ax.set_ylabel("USD")
ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f"${x:,.0f}"))
plt.xticks(rotation=30)
plt.tight_layout()
plt.savefig("figures/03_daily_trend.png", dpi=140, bbox_inches="tight")
plt.show()
""")

code("""
# Weekday pattern — when do I spend?
wd_order = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]
wd = (df.groupby("Weekday")["Amount"]
        .agg(Total="sum", Count="count")
        .reindex(wd_order)
        .round(2))
wd["AvgPerDay"] = (wd["Total"] / wd["Count"]).round(2)

fig, ax = plt.subplots(figsize=(10, 5))
bars = ax.bar(wd.index, wd["Total"], color=["#1F4E78"]*5 + ["#F79256"]*2)
ax.set_title("Spend by Day of Week  (weekends highlighted)")
ax.set_ylabel("USD")
ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f"${x:,.0f}"))
for b, v in zip(bars, wd["Total"]):
    ax.text(b.get_x()+b.get_width()/2, v, f"${v:,.0f}", ha="center", va="bottom", fontsize=10)
plt.xticks(rotation=20)
plt.tight_layout()
plt.savefig("figures/04_weekday_pattern.png", dpi=140, bbox_inches="tight")
plt.show()
wd
""")

code("""
# Payment-method breakdown
pm = (df.groupby("Payment Method")["Amount"]
        .agg(Total="sum", Count="count")
        .sort_values("Total", ascending=False))
pm["Share"] = (pm["Total"] / pm["Total"].sum()).map("{:.1%}".format)
pm["AvgPerTxn"] = (pm["Total"] / pm["Count"]).round(2)

fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12, 5))
colors = ["#2E75B6", "#F79256"]
ax1.pie(pm["Total"], labels=pm.index, autopct="%1.1f%%", colors=colors, startangle=90,
        wedgeprops=dict(edgecolor="white", linewidth=2))
ax1.set_title("Share of Spending by Payment Method")

pm["Total"].plot.bar(ax=ax2, color=colors)
ax2.set_title("Total Spend by Payment Method")
ax2.set_ylabel("USD")
ax2.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f"${x:,.0f}"))
ax2.tick_params(axis="x", rotation=0)
for i, v in enumerate(pm["Total"]):
    ax2.text(i, v, f"${v:,.0f}", ha="center", va="bottom", fontsize=11, fontweight="bold")
plt.tight_layout()
plt.savefig("figures/05_payment_methods.png", dpi=140, bbox_inches="tight")
plt.show()
pm
""")

code("""
# Top 10 descriptions (merchant-ish) by total spend
top_merch = (df.groupby("Description")["Amount"]
               .agg(Total="sum", Visits="count")
               .sort_values("Total", ascending=False)
               .head(10)
               .round(2))
top_merch
""")

md("## 4 · Model — month-over-month change and category mix")

code("""
# Month-over-month change in total spend
mom = monthly.copy()
mom["ΔFromPrev"] = mom["Spend"].diff().round(2)
mom["% Change"]  = (mom["Spend"].pct_change() * 100).round(1)
mom
""")

code("""
# Category × Month pivot — where did spend shift?
pivot = (df.pivot_table(index="Category", columns="YearMonth",
                        values="Amount", aggfunc="sum", fill_value=0)
           .round(2))
# Focus on the largest categories for readability
top_n = cat_totals.head(10).index
pivot_top = pivot.loc[top_n]

fig, ax = plt.subplots(figsize=(11, 7))
sns.heatmap(pivot_top, cmap="Blues", annot=True, fmt=".0f",
            cbar_kws={"label": "USD"}, ax=ax,
            linewidths=.3, linecolor="white")
ax.set_title("Top 10 Categories × Month")
ax.set_xlabel("Month")
ax.set_ylabel("")
plt.tight_layout()
plt.savefig("figures/06_category_month_heatmap.png", dpi=140, bbox_inches="tight")
plt.show()
""")

code("""
# Simple linear trend on monthly spend (with only 4 points, this is
# directional — not forecastable with confidence)
x = np.arange(len(monthly))
y = monthly["Spend"].values
slope, intercept = np.polyfit(x, y, 1)
trend = intercept + slope * x

print(f"Linear fit on 4 monthly points:")
print(f"  Slope: ${slope:,.0f}/month  (negative = spending coming down)")
print(f"  Intercept: ${intercept:,.0f}")
print(f"  Caveat: n=4 is too small for a real forecast; treat as directional only.")

fig, ax = plt.subplots(figsize=(10, 5))
ax.plot(monthly["YearMonth"], y, marker="o", markersize=10, color="#1F4E78", label="Actual", linewidth=2)
ax.plot(monthly["YearMonth"], trend, linestyle="--", color="#c0392b", label="Linear trend")
ax.set_title("Spend Trend — Linear Fit (n=4, directional)")
ax.set_ylabel("USD")
ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f"${x:,.0f}"))
ax.legend()
plt.tight_layout()
plt.savefig("figures/07_linear_trend.png", dpi=140, bbox_inches="tight")
plt.show()
""")

code("""
# Simplified Needs / Wants bucket mapping, using the CONSOLIDATED categories
# (clean_data.py has already merged synonyms like Supermarkets→Groceries).
needs_cats = ["Housing", "Utilities", "Groceries", "Transport",
              "Gasoline", "Government", "Services",
              "Personal Care", "Fines/Legal", "Subscriptions"]
wants_cats = ["Restaurants", "Coffee", "Entertainment", "Shopping",
              "Office Supplies"]

df["Bucket"] = np.select(
    [df["Category"].isin(needs_cats), df["Category"].isin(wants_cats)],
    ["Needs", "Wants"],
    default="Other",
)
bucket = df.groupby("Bucket")["Amount"].sum().round(2)
bucket_share = (bucket / bucket.sum() * 100).round(1)
print("Bucket split:")
print(pd.concat([bucket.rename("USD"),
                 bucket_share.rename("% of spend")], axis=1))
print("\\n50/30/20 rule target: Needs 50% · Wants 30% · Savings 20%")
print("(Savings isn't in the expense feed — it'd need to come from income data.)")
""")

md("## 5 · iNterpret — key findings & recommendations")

code("""
# Pull the numbers we want to reference in the write-up
top_cat_name = cat_totals.index[0]
top_cat_val  = cat_totals["Total"].iloc[0]
top_cat_pct  = top_cat_val / cat_totals["Total"].sum() * 100

restaurants  = df[df["Category"] == "Restaurants"]["Amount"].sum()
coffee       = df[df["Category"] == "Coffee"]["Amount"].sum()
dining_total = restaurants + coffee

bofa_share   = pm.loc["BofA Credit", "Total"] / pm["Total"].sum() * 100
peak_month   = monthly.loc[monthly["Spend"].idxmax(), "YearMonth"]
peak_val     = monthly["Spend"].max()

print(f"Total spend:               ${df['Amount'].sum():,.2f}")
print(f"Top category:              {top_cat_name} — ${top_cat_val:,.2f} ({top_cat_pct:.1f}%)")
print(f"Restaurants + Coffee:      ${dining_total:,.2f}  ({dining_total/df['Amount'].sum()*100:.1f}% of total)")
print(f"BofA Credit share:         {bofa_share:.1f}%  ({int(pm.loc['BofA Credit','Count'])} of {len(df)} transactions)")
print(f"Peak month:                {peak_month} — ${peak_val:,.2f}")
""")

md("""### Findings

1. **Restaurants is the single biggest line item** — it alone accounts for ~15%
   of total spend over this 3.5-month window, and Restaurants + Coffee combined
   push that over 17%. This is the most actionable lever.
2. **January was the peak month** (~$2.2k) — that's when the USCIS I-765 fee
   ($470), a Mint Mobile renewal ($267), and the Summit One Vanderbilt visit
   landed in the same month. Pulling those out, the baseline is closer to
   ~$1.1–1.4k/month.
3. **BofA Credit carries the vast majority of transactions** (~76% of spend,
   76% of all swipes) — Discover is used for a much smaller share, mostly for
   specific merchants.
4. **Weekday pattern skews toward weekends** — Saturday + Sunday together
   account for a meaningful chunk of dining-out and shopping trips.
5. **One-offs distort the month-over-month view.** A handful of large,
   non-recurring charges (USCIS, Eversource, Summit One Vanderbilt) swing any
   given month by hundreds of dollars — recurring spend is smoother than the
   totals suggest.

### Recommendations

- Set a **weekly Restaurants + Coffee target** (e.g., cap at 75% of the current
  weekly average) — this is the highest-leverage category.
- **Separate one-offs from recurring** in the monthly view (e.g., a "discretionary
  recurring" KPI that excludes government fees, utility bills, and annual
  subscriptions) to get a cleaner baseline.
- Revisit **recurring subscriptions** (Mint Mobile, etc.) once per quarter — small
  monthly charges add up; annual renewals like the $267 Mint Mobile charge
  should be intentional, not passive.
- Consider **routing more recurring bills through Discover** (or whichever card
  has the better category rewards) to balance utilization and rewards.

Next step → the companion Excel dashboard turns these insights into an
interactive tool you can re-slice by Month / Category / Payment Method.
""")

# ---------- execute cells in a single namespace & capture outputs ----------
Path(HERE / "figures").mkdir(exist_ok=True)
ns: dict = {"__name__": "__main__"}

def run_code(src):
    """Execute src; return list of output dicts (stream/text/image)."""
    outputs = []
    buf = io.StringIO()

    plt.close("all")

    last = {"val": None}
    class DisplayHook:
        def __init__(self, old): self.old = old
        def __call__(self, value):
            if value is None: return
            last["val"] = value
    old_hook = sys.displayhook
    sys.displayhook = DisplayHook(old_hook)
    try:
        with contextlib.redirect_stdout(buf):
            try:
                import ast
                mod = ast.parse(src, mode="exec")
                if mod.body:
                    *body, tail = mod.body
                    if body:
                        exec(compile(ast.Module(body=body, type_ignores=[]), "<cell>", "exec"), ns)
                    if isinstance(tail, ast.Expr):
                        exec(compile(ast.Interactive(body=[tail]), "<cell>", "single"), ns)
                    else:
                        exec(compile(ast.Module(body=[tail], type_ignores=[]), "<cell>", "exec"), ns)
            except Exception:
                traceback.print_exc()
    finally:
        sys.displayhook = old_hook

    text = buf.getvalue()
    if text.strip():
        outputs.append({
            "output_type": "stream",
            "name": "stdout",
            "text": text.splitlines(keepends=True),
        })

    val = last["val"]
    if val is not None:
        data = {}
        try:
            import pandas as pd
            if isinstance(val, (pd.DataFrame, pd.Series)):
                data["text/html"]  = val.to_html() if isinstance(val, pd.DataFrame) else val.to_frame().to_html()
                data["text/plain"] = repr(val)
        except Exception:
            pass
        if not data:
            data["text/plain"] = repr(val)
        outputs.append({
            "output_type": "execute_result",
            "execution_count": None,
            "data": data,
            "metadata": {},
        })

    fignums = plt.get_fignums()
    for num in fignums:
        fig = plt.figure(num)
        img_buf = io.BytesIO()
        fig.savefig(img_buf, format="png", dpi=120, bbox_inches="tight")
        img_buf.seek(0)
        outputs.append({
            "output_type": "display_data",
            "data": {"image/png": base64.b64encode(img_buf.read()).decode("ascii")},
            "metadata": {"image/png": {"width": 720}},
        })
        plt.close(fig)

    return outputs


cells_json = []
exec_counter = 0
for kind, src in C:
    if kind == "md":
        cells_json.append({
            "cell_type": "markdown",
            "metadata": {},
            "source": src.splitlines(keepends=True),
        })
    else:
        outs = run_code(src)
        exec_counter += 1
        cells_json.append({
            "cell_type": "code",
            "execution_count": exec_counter,
            "metadata": {},
            "outputs": outs,
            "source": src.splitlines(keepends=True),
        })

nb = {
    "cells": cells_json,
    "metadata": {
        "kernelspec":    {"display_name": "Python 3", "language": "python", "name": "python3"},
        "language_info": {"name": "python", "version": "3.10"},
        "title": "Personal Finance — OSEMN Analysis (Real Data)",
        "created": datetime.datetime.utcnow().isoformat() + "Z",
    },
    "nbformat": 4,
    "nbformat_minor": 5,
}
OUT_NB.write_text(json.dumps(nb, indent=1))
print(f"Wrote notebook ({OUT_NB.stat().st_size/1024:.1f} KB): {OUT_NB}")
print(f"Figures in: {HERE/'figures'}")
