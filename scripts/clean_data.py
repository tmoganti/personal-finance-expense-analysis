"""
OSEMN Step 1 & 2: Obtain the real Expenses.numbers file (converted to CSV by
LibreOffice) and scrub it into a clean, analysis-ready expenses.csv.

Changes made:
  • Strip trailing spaces from column headers.
  • Fix 4 rows dated "2026-12-31" → "2025-12-31" (obvious typo: they sit
    between 2025-12-30 and 2026-01-01 entries).
  • Parse Date as ISO dates.
  • Consolidate overlapping categories (Supermarkets→Groceries,
    Gas/Convenience + Convenience Store→Gasoline, Merchandise→Shopping,
    Online Services→Subscriptions) so per-category analyses aren't split
    across synonyms. The original label is kept in Category_Raw for audit.
  • Flag non-recurring one-off transactions (Government fees, Fines/Legal,
    big single-visit experiences) with IsOneOff = "Yes"/"No" so downstream
    analysis can strip them to show the recurring baseline.
  • Derive Year, Month, YearMonth, Weekday columns.
  • Assign a synthetic TransactionID for traceability.
"""
from pathlib import Path
import pandas as pd

HERE = Path(__file__).resolve().parent
SRC  = Path("/sessions/bold-peaceful-clarke/mnt/outputs/Expenses.csv")
OUT  = HERE / "expenses.csv"

# ---------- category consolidation mapping ----------
CATEGORY_MAP = {
    "Supermarkets":        "Groceries",
    "Gas/Convenience":     "Gasoline",
    "Convenience Store":   "Gasoline",
    "Merchandise":         "Shopping",
    "Online Services":     "Subscriptions",
}

# ---------- one-off transaction rules ----------
ONE_OFF_CATEGORIES = {"Government", "Fines/Legal"}
ONE_OFF_DESCRIPTIONS = {
    "Mint Mobile Subscription",           # annual renewal — not monthly
    "Summit One Vanderbilt NY",           # tourist visit
    "Abercrombie & Fitch Burlington MA",  # one-time shopping trip
    "Eversource Utility Bill",            # bimonthly lump, distorts monthly view
}

# ============================================================
df = pd.read_csv(SRC)
df.columns = [c.strip() for c in df.columns]

# Fix typo: 2026-12-31 rows should be 2025-12-31 (only Dec entries before Jan 01)
mask = df["Date"] == "2026-12-31"
assert mask.sum() == 4, f"Expected 4 typo rows, found {mask.sum()}"
df.loc[mask, "Date"] = "2025-12-31"

df["Date"]      = pd.to_datetime(df["Date"])
df["Year"]      = df["Date"].dt.year
df["Month"]     = df["Date"].dt.month
df["YearMonth"] = df["Date"].dt.strftime("%Y-%m")
df["Weekday"]   = df["Date"].dt.day_name()

df["Category_Raw"] = df["Category"]
df["Category"] = df["Category"].replace(CATEGORY_MAP)

df["IsOneOff"] = (
    df["Category"].isin(ONE_OFF_CATEGORIES)
    | df["Description"].isin(ONE_OFF_DESCRIPTIONS)
).map({True: "Yes", False: "No"})

df = df.sort_values("Date").reset_index(drop=True)
df["TransactionID"] = ["T{:04d}".format(i + 1) for i in range(len(df))]

df = df[["TransactionID", "Date", "Year", "Month", "YearMonth",
         "Description", "Category", "Category_Raw",
         "Amount", "Payment Method", "Weekday", "IsOneOff"]]

df.to_csv(OUT, index=False)

print(f"Saved {len(df)} rows → {OUT}")
print(f"Date range: {df['Date'].min().date()} → {df['Date'].max().date()}")
print(f"Total spend: ${df['Amount'].sum():,.2f}")

print(f"\nCategory consolidation (rows remapped):")
remap = (df[df['Category'] != df['Category_Raw']]
         .groupby(['Category_Raw', 'Category'])
         .size().reset_index(name='rows'))
print(remap.to_string(index=False) if len(remap) else "  (none)")

one_off_total = df.loc[df['IsOneOff']=='Yes','Amount'].sum()
print(f"\nOne-off transactions: {(df['IsOneOff']=='Yes').sum()} rows, "
      f"${one_off_total:,.2f} ({one_off_total/df['Amount'].sum()*100:.1f}% of total)")
print("One-off breakdown:")
print(df.loc[df['IsOneOff']=='Yes']
        .groupby(['Category','Description'])['Amount']
        .agg(['sum','count']).round(2))

print(f"\nConsolidated category totals:")
print(df.groupby('Category')['Amount'].sum().sort_values(ascending=False).round(2))

print(f"\nMonthly totals — All vs Recurring:")
monthly = df.groupby('YearMonth')['Amount'].sum()
recur   = df.loc[df['IsOneOff']=='No'].groupby('YearMonth')['Amount'].sum()
print(pd.DataFrame({'All': monthly, 'Recurring only': recur}).round(2))
