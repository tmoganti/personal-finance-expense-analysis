"""
Render Dashboard_Preview.png — a single 1920×1080 composite that mirrors
the interactive dashboard in a format you can post to LinkedIn, drop in
slide decks, or thumbnail in a portfolio README.

Layout (navy/blue palette, matches the Excel dashboard):

  ┌──────────────────────────────────────────────────────────────┐
  │           PERSONAL FINANCE DASHBOARD · Dec 2025 – Mar 2026   │
  ├──────────────────────────────────────────────────────────────┤
  │ TOTAL │ MONTHLY │ TRANS. │ TOP CAT. │ ONE-OFF │ RECURRING     │
  ├────────────────────┬─────────────────────────────────────────┤
  │ Spend by Category  │  Monthly — Total vs Recurring           │
  ├────────────────────┼───────────────────┬─────────────────────┤
  │ Top 10 Merchants   │ Payment Method    │ Daily Spend         │
  ├────────────────────┴───────────────────┴─────────────────────┤
  │ KEY INSIGHTS                                                 │
  └──────────────────────────────────────────────────────────────┘
"""
from pathlib import Path
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from matplotlib.gridspec import GridSpec

# Disable math-text parsing globally so "$" in labels renders as a literal dollar
plt.rcParams["text.usetex"] = False
plt.rcParams["mathtext.default"] = "regular"

ROOT = Path(__file__).resolve().parents[1]
CSV  = ROOT / "expenses.csv"
OUT  = ROOT / "Dashboard_Preview.png"

# Palette mirrors spending_dashboard.html — cream background, ink-black title
# band, terracotta accent. Keeps the README hero aesthetically aligned with the
# live HTML deliverable.
NAVY    = "#1a3a6e"   # deep blue, used for "Total" series
ACCENT  = "#c8522a"   # terracotta, used for "Recurring" series
ACCENT2 = "#1a9e6e"   # teal, used for highlights
INK     = "#1a1814"   # near-black for title band & body text
SURFACE = "#ffffff"   # KPI card surface
BG      = "#f5f2ec"   # page background
LIGHT   = "#f0ede6"   # subtle card border / band fill
GRAY    = "#6b6760"
WHITE   = "#ffffff"
# Legacy aliases so the rest of the script reads naturally.
BLUE, SKY = NAVY, "#dfe6f0"

df = pd.read_csv(CSV, parse_dates=["Date"])

# Aggregates
total = df["Amount"].sum()
n_months = df["YearMonth"].nunique()
n_txn = len(df)
cat_sum = df.groupby("Category")["Amount"].sum().sort_values(ascending=False)
top_cat = cat_sum.index[0]
one_off_total = df.loc[df["IsOneOff"] == "Yes", "Amount"].sum()
recur_total   = total - one_off_total

monthly = df.groupby("YearMonth")["Amount"].sum()
monthly_recur = df[df["IsOneOff"] == "No"].groupby("YearMonth")["Amount"].sum()

pm = df.groupby("Payment Method")["Amount"].sum().sort_values(ascending=False)
merch = df.groupby("Description")["Amount"].sum().sort_values(ascending=True).tail(10)
daily = df.groupby("Date")["Amount"].sum()

# -------------------------------------------------- figure
fig = plt.figure(figsize=(19.2, 10.8), dpi=100, facecolor=WHITE)
gs = GridSpec(
    nrows=22, ncols=12, figure=fig,
    left=0.025, right=0.975, top=0.97, bottom=0.03,
    hspace=1.0, wspace=0.6,
)

# ---- Title band (rows 0–1)
ax_t = fig.add_subplot(gs[0:2, :]); ax_t.axis("off")
ax_t.add_patch(patches.Rectangle((0, 0), 1, 1, transform=ax_t.transAxes,
                                  facecolor=NAVY, edgecolor="none"))
ax_t.text(0.5, 0.55, "PERSONAL FINANCE DASHBOARD",
          transform=ax_t.transAxes, ha="center", va="center",
          color=WHITE, fontsize=28, fontweight="bold")
ax_t.text(0.5, 0.15,
          f"{df['Date'].min():%b %Y}  →  {df['Date'].max():%b %Y}   ·   "
          f"{n_txn} transactions   ·   ${total:,.2f} total",
          transform=ax_t.transAxes, ha="center", va="center",
          color=LIGHT, fontsize=14, style="italic")

# ---- KPI band (rows 2–4) — 6 cards across 12 cols (each 2 cols wide)
def kpi(col_start, label, value, fmt, sub=""):
    ax = fig.add_subplot(gs[2:4, col_start:col_start + 2]); ax.axis("off")
    ax.add_patch(patches.FancyBboxPatch(
        (0.02, 0.05), 0.96, 0.90,
        boxstyle="round,pad=0.02",
        transform=ax.transAxes, facecolor=LIGHT,
        edgecolor=NAVY, linewidth=1.2))
    ax.text(0.5, 0.78, label, transform=ax.transAxes,
            ha="center", va="center", color=NAVY, fontsize=10.5, fontweight="bold")
    ax.text(0.5, 0.43, fmt.format(value), transform=ax.transAxes,
            ha="center", va="center", color=NAVY, fontsize=22, fontweight="bold")
    if sub:
        ax.text(0.5, 0.15, sub, transform=ax.transAxes,
                ha="center", va="center", color=GRAY, fontsize=9.5)

kpi(0,  "TOTAL SPEND",    total,                   "${:,.0f}")
kpi(2,  "MONTHLY AVG",    total / n_months,        "${:,.0f}", f"across {n_months} months")
kpi(4,  "TRANSACTIONS",   n_txn,                   "{:,}")
kpi(6,  "TOP CATEGORY",   top_cat,                 "{}",        f"${cat_sum.iloc[0]:,.0f}")
kpi(8,  "ONE-OFF CHARGES", one_off_total,           "${:,.0f}", f"{one_off_total/total*100:.0f}% of total")
kpi(10, "RECURRING",       recur_total,             "${:,.0f}", "after stripping one-offs")

# ---- Row 1 of charts: Spend by Category (left wide)  |  Monthly Total vs Recurring (right wide)
ax1 = fig.add_subplot(gs[5:12, 0:6])
cat_sum_sorted = cat_sum.sort_values(ascending=True)
ax1.barh(cat_sum_sorted.index, cat_sum_sorted.values, color=BLUE)
ax1.set_title("Spend by Category", loc="left", fontsize=13, fontweight="bold", color=NAVY, pad=8)
ax1.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f"${x:,.0f}"))
ax1.spines["top"].set_visible(False); ax1.spines["right"].set_visible(False)
ax1.tick_params(axis="both", labelsize=9)
for i, v in enumerate(cat_sum_sorted.values):
    ax1.text(v, i, f"  ${v:,.0f}", va="center", fontsize=8.5, color=INK)

ax2 = fig.add_subplot(gs[5:12, 6:12])
x = np.arange(len(monthly))
w = 0.38
ax2.bar(x - w/2, monthly.values,        w, label="Total",          color=NAVY)
ax2.bar(x + w/2, monthly_recur.reindex(monthly.index).values, w,
        label="Recurring only", color=ACCENT)
ax2.set_xticks(x); ax2.set_xticklabels(monthly.index, fontsize=10)
ax2.set_title("Monthly Spend — Total vs Recurring", loc="left",
              fontsize=13, fontweight="bold", color=NAVY, pad=8)
ax2.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f"${x:,.0f}"))
ax2.spines["top"].set_visible(False); ax2.spines["right"].set_visible(False)
ax2.legend(loc="upper right", fontsize=10)
for i, v in enumerate(monthly.values):
    ax2.text(i - w/2, v, f"${v:,.0f}", ha="center", va="bottom",
             fontsize=9, fontweight="bold", color=NAVY)
for i, v in enumerate(monthly_recur.reindex(monthly.index).values):
    ax2.text(i + w/2, v, f"${v:,.0f}", ha="center", va="bottom",
             fontsize=9, fontweight="bold", color=ACCENT)
ax2.set_ylim(0, monthly.max() * 1.18)

# ---- Row 2 of charts: Top Merchants | Payment Method | Daily Spend
ax3 = fig.add_subplot(gs[13:20, 0:4])
ax3.barh(merch.index, merch.values, color=ACCENT)
ax3.set_title("Top 10 Merchants", loc="left", fontsize=13, fontweight="bold", color=NAVY, pad=8)
ax3.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f"${x:,.0f}"))
ax3.spines["top"].set_visible(False); ax3.spines["right"].set_visible(False)
ax3.tick_params(axis="both", labelsize=8.5)

ax4 = fig.add_subplot(gs[13:20, 4:7])
colors = [BLUE, ACCENT, SKY, GRAY][:len(pm)]
wedges, texts, autotexts = ax4.pie(
    pm.values, labels=pm.index, colors=colors, autopct="%1.1f%%",
    startangle=90, textprops=dict(fontsize=10, color=INK),
    wedgeprops=dict(edgecolor=WHITE, linewidth=2),
)
for a in autotexts:
    a.set_fontweight("bold"); a.set_color(WHITE); a.set_fontsize(11)
ax4.set_title("Payment Method Share", loc="left",
              fontsize=13, fontweight="bold", color=NAVY, pad=8)

ax5 = fig.add_subplot(gs[13:20, 7:12])
ax5.plot(daily.index, daily.values, marker="o", markersize=4,
         color=BLUE, linewidth=1.3)
ax5.fill_between(daily.index, daily.values, alpha=0.25, color=SKY)
ax5.set_title("Daily Spend", loc="left",
              fontsize=13, fontweight="bold", color=NAVY, pad=8)
ax5.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f"${x:,.0f}"))
ax5.spines["top"].set_visible(False); ax5.spines["right"].set_visible(False)
ax5.tick_params(axis="both", labelsize=9)
fig.autofmt_xdate()

# ---- Insights band (row 20–21)
top_merch_name = df.groupby("Description")["Amount"].sum().idxmax()
top_merch_amt  = df.groupby("Description")["Amount"].sum().max()
bofa_share     = pm.get("BofA Credit", 0) / pm.sum() * 100
recur_top = (df[df["IsOneOff"]=="No"].groupby("Category")["Amount"]
             .sum().sort_values(ascending=False))
insight_text = (
    f"KEY INSIGHTS     "
    f"•  One-offs: \${one_off_total:,.0f} ({one_off_total/total*100:.0f}%) — "
    f"recurring baseline = \${recur_total:,.0f}    "
    f"•  Recurring top category: {recur_top.index[0]} (\${recur_top.iloc[0]:,.0f})    "
    f"•  BofA carries {bofa_share:.0f}% of spend    "
    f"•  Biggest merchant: {top_merch_name} (\${top_merch_amt:,.0f})"
)
ax_i = fig.add_subplot(gs[20:22, :]); ax_i.axis("off")
ax_i.add_patch(patches.Rectangle((0, 0), 1, 1, transform=ax_i.transAxes,
                                  facecolor=NAVY, edgecolor="none"))
ax_i.text(0.5, 0.5, insight_text, transform=ax_i.transAxes,
          ha="center", va="center", color=WHITE, fontsize=12)

fig.savefig(OUT, dpi=100, bbox_inches=None, facecolor=WHITE)
plt.close(fig)
print(f"Saved {OUT}  ({OUT.stat().st_size/1024:.1f} KB)")
