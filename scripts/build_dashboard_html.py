"""
build_dashboard_html.py — Regenerate spending_dashboard.html data block from expenses.csv.

The HTML template `spending_dashboard.html` ships with two markers in its
<script> block:

    // __DATA_START__
    const ALL_TXN = [ ... ];
    // __DATA_END__

This script reads `expenses.csv`, formats every row into a JS object literal,
and rewrites everything between those markers. The HTML's styling, charts,
and Chart.js logic are untouched — only the underlying data is refreshed.

Pipeline shape (mirrors build_dashboard.py for Excel):

    expenses.csv ──► build_dashboard_html.py ──► spending_dashboard.html

Run:
    python scripts/build_dashboard_html.py

Idempotent — safe to re-run after every clean_data.py update.
"""

from __future__ import annotations

import csv
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
CSV_PATH = ROOT / "expenses.csv"
HTML_PATH = ROOT / "spending_dashboard.html"

# ── YearMonth → tab key in the dashboard ──────────────────────────────
MONTH_KEY = {
    "2025-12": "dec",
    "2026-01": "jan",
    "2026-02": "feb",
    "2026-03": "mar",
}

# ── Payment Method → card key in the dashboard ────────────────────────
CARD_KEY = {
    "Discover Card": "discover",
    "BofA Credit": "bofa",
}

# Strip city/state suffixes from merchant names so they fit the merchant table.
# Order matters — apply more specific patterns first.
_MERCHANT_CLEANUPS = [
    (re.compile(r"\s+(?:Online|Subscription|Online Services|US Online Store)$", re.I), ""),
    (re.compile(r"\s+(?:Greater\s+)?Nashua\s+NH$", re.I), ""),
    (re.compile(r"\s+(?:Burlington|Cambridge|Methuen|Lyndhurst|Salem|North Billerica|Tyngsboro|Boston|Lowell|Hudson|Cranbury|Farmington|Wilmington|Pay\s*Go|PayGo)\s+(?:NH|MA|NJ|CT|DE|NY)$", re.I), ""),
    (re.compile(r"\s+New\s+York\s+NY$", re.I), " NYC"),
    (re.compile(r"\s+II\s+Cambridge\s+MA$", re.I), ""),
    (re.compile(r"^MTA\s+NYCT\s+PayGo.*", re.I), "MTA NYC"),
    (re.compile(r"^Logan\s+Airport\s+Parking.*", re.I), "Logan Parking"),
    (re.compile(r"^Vpdfhouse\s+(Online Services|Wilmington DE)$", re.I), "Vpdfhouse"),
    (re.compile(r"^The\s+Jewelers\s+Workbench.*", re.I), "Jewelers Workbench"),
    (re.compile(r"^Macys\s+Starbucks.*", re.I), "Starbucks"),
    (re.compile(r"^Schoharie\s+Town\s+Court\s+NY$", re.I), "Schoharie Town Court"),
    (re.compile(r"^Schoharie\s+Court\s+Fee$", re.I), "Schoharie Court Fee"),
]

def clean_merchant(raw: str) -> str:
    """Trim location suffixes so 'Patel Brothers Nashua NH' becomes 'Patel Brothers'."""
    name = raw.strip()
    for pattern, replacement in _MERCHANT_CLEANUPS:
        name = pattern.sub(replacement, name)
    return name.strip()


def js_string(s: str) -> str:
    """Quote a string for safe embedding in a JS object literal."""
    escaped = s.replace("\\", "\\\\").replace("'", "\\'")
    return f"'{escaped}'"


def row_to_js(row: dict) -> str:
    m = MONTH_KEY.get(row["YearMonth"])
    card = CARD_KEY.get(row["Payment Method"])
    if m is None or card is None:
        raise ValueError(
            f"Unknown YearMonth or Payment Method on row {row.get('TransactionID')}: "
            f"{row['YearMonth']!r} / {row['Payment Method']!r}"
        )
    merchant = clean_merchant(row["Description"])
    cat = row["Category"]
    amt = float(row["Amount"])
    parts = [
        f"m:'{m}'",
        f"card:'{card}'",
        f"merchant:{js_string(merchant)}",
        f"cat:{js_string(cat)}",
        f"amt:{amt:.2f}",
    ]
    if row.get("IsOneOff", "").strip().lower() == "yes":
        parts.append("oneoff:true")
    return "  {" + ",".join(parts) + "},"


def build_data_block(rows: list[dict]) -> str:
    """Render the const ALL_TXN = [ ... ]; block, grouped by month with section headers."""
    by_month: dict[str, list[dict]] = {"dec": [], "jan": [], "feb": [], "mar": []}
    for row in rows:
        key = MONTH_KEY.get(row["YearMonth"])
        if key in by_month:
            by_month[key].append(row)

    section_titles = {
        "dec": "December 2025",
        "jan": "January 2026  (one-offs flagged)",
        "feb": "February 2026",
        "mar": "March 2026",
    }

    lines = [
        "// __DATA_START__",
        "const ALL_TXN = [",
    ]
    for key in ["dec", "jan", "feb", "mar"]:
        lines.append(f"  // ── {section_titles[key]} ──")
        for row in sorted(by_month[key], key=lambda r: (r["Date"], r["TransactionID"])):
            lines.append(row_to_js(row))
    lines.append("];")
    lines.append("// __DATA_END__")
    return "\n".join(lines)


def main() -> None:
    if not CSV_PATH.exists():
        raise SystemExit(f"expenses.csv not found at {CSV_PATH}")
    if not HTML_PATH.exists():
        raise SystemExit(f"spending_dashboard.html not found at {HTML_PATH}")

    with CSV_PATH.open(newline="", encoding="utf-8") as fh:
        rows = list(csv.DictReader(fh))

    new_block = build_data_block(rows)

    html = HTML_PATH.read_text(encoding="utf-8")
    pattern = re.compile(r"// __DATA_START__.*?// __DATA_END__", re.DOTALL)
    if not pattern.search(html):
        raise SystemExit(
            "Could not find // __DATA_START__ ... // __DATA_END__ markers in "
            "spending_dashboard.html. The template was edited — restore the "
            "markers around the const ALL_TXN = [...] block before re-running."
        )
    new_html = pattern.sub(new_block, html, count=1)
    HTML_PATH.write_text(new_html, encoding="utf-8")

    total = sum(float(r["Amount"]) for r in rows)
    oneoffs = [r for r in rows if r.get("IsOneOff", "").lower() == "yes"]
    oneoff_total = sum(float(r["Amount"]) for r in oneoffs)
    print(f"Wrote {HTML_PATH.name}")
    print(f"  rows:        {len(rows)}")
    print(f"  total spend: ${total:,.2f}")
    print(f"  one-offs:    {len(oneoffs)} txns, ${oneoff_total:,.2f} ({oneoff_total/total*100:.1f}%)")


if __name__ == "__main__":
    main()
