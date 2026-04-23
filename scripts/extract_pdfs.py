"""
OSEMN Step 1 (Obtain) — real bank/credit-card statement ingestion.

Walks the workspace folder, detects the statement type, and writes a
unified CSV that matches the downstream `expenses.csv` schema used by
clean_data.py.

Supported formats
-----------------
1. Discover It statements  ->  "January 2026.pdf", "February 2026.pdf", "March 2026.pdf"
   Format: page 3 contains the PURCHASES section where each row is
     MM/DD <Description>  <MerchantCategory>  $<Amount>
   Categories are provided by the issuer.

2. BofA Visa Signature credit card statements  ->  eStmt_2026-01-27.pdf,
   eStmt_2026-02-27.pdf, eStmt_2026-03-27.pdf  (statements dated the
   27th-ish of the month).
   Format: page 3+ has a "Purchases and Adjustments" block with rows
     TRANS_DATE  POST_DATE  Description  RefNum  AcctNum  Amount
   No category is given; set Category = "Uncategorized" so the analyst
   can backfill (e.g. via a manual merchant→category lookup).

3. BofA SafeBalance Checking statements  ->  eStmt_2026-01-23.pdf etc.
   These are deposit + debit-card account statements with a very
   different layout (deposits vs. withdrawals vs. card purchases). They
   are intentionally *not* parsed here — a future iteration can add a
   dedicated parser. We log and skip them so nothing is silently lost.

Output
------
  extracted_from_pdfs.csv      — one row per purchase, same schema as expenses.csv
                                 (pre-clean_data: Date, Description, Category,
                                 Amount, Payment Method)

This file feeds the same downstream pipeline (clean_data.py can concat
it with any manually-tracked CSV, or it can be used standalone).
"""
from pathlib import Path
import re
import pdfplumber
import pandas as pd

# ------------------------------------------------------------------
# Paths
# ------------------------------------------------------------------
WORKSPACE = Path("/sessions/bold-peaceful-clarke/mnt/Personal Finance and Expense Analysis")
HERE      = Path(__file__).resolve().parent
OUT_CSV   = HERE / "extracted_from_pdfs.csv"

# ------------------------------------------------------------------
# Helpers
# ------------------------------------------------------------------
MONTH_RANGE_RX = re.compile(
    r"OPEN TO CLOSE DATE:\s*(\d{2}/\d{2}/\d{4})\s*-\s*(\d{2}/\d{2}/\d{4})"
)
BOFA_PERIOD_RX = re.compile(
    r"(January|February|March|April|May|June|July|August|September|October|November|December)"
    r"\s+(\d{1,2})\s*-\s*"
    r"(January|February|March|April|May|June|July|August|September|October|November|December)?"
    r"\s*(\d{1,2}),\s*(\d{4})"
)

DISCOVER_KNOWN_CATS = [
    # Order matters: longer names first so "Travel/Entertainment" wins over "Travel"
    "Travel/Entertainment", "Government Services", "Department Stores",
    "Home Improvement", "Medical Services",
    "Merchandise", "Restaurants", "Supermarkets", "Gasoline",
    "Services", "Education", "Automotive",
]

# Discover transaction row (date + body + $amount; Discover often prints
# sidebar advertising text after the amount on the same line):
#   12/12 ADIDAS US ONLINE STORE 800-982-9337 Merchandise $105.00 Streaming Services
DISCOVER_TXN_RX = re.compile(
    r"^\s*(\d{2}/\d{2})\s+(.+?)\s+\$([\d,]+\.\d{2})(?:\s|$)"
)

# BofA Visa credit-card transaction row:
#   12/31 12/31 OPENAI *CHATGPT SUBSCR OPENAI.COM CA 3040 6229 20.00
BOFA_TXN_RX = re.compile(
    r"^\s*(\d{2}/\d{2})\s+(\d{2}/\d{2})\s+(.+?)\s+(\d{4})\s+(\d{4})\s+([\-\d,]+\.\d{2})\s*$"
)


def _infer_year(mm_dd: str, statement_start: pd.Timestamp, statement_end: pd.Timestamp) -> int:
    """A statement period that crosses a year boundary (e.g. 12/06/2025 → 01/05/2026)
    means some MM/DD rows belong to the start year and others to the end year.
    Pick whichever makes the date fall inside the [start, end] window.
    """
    mm, dd = map(int, mm_dd.split("/"))
    for year in (statement_start.year, statement_end.year):
        candidate = pd.Timestamp(year=year, month=mm, day=dd)
        if statement_start <= candidate <= statement_end:
            return year
    # Fallback — shouldn't normally hit this
    return statement_end.year


# ------------------------------------------------------------------
# Parser: Discover
# ------------------------------------------------------------------
def parse_discover(pdf_path: Path) -> list[dict]:
    rows: list[dict] = []
    with pdfplumber.open(pdf_path) as pdf:
        # Figure out the statement window from page 3's header
        header_txt = pdf.pages[2].extract_text() or ""
        m = MONTH_RANGE_RX.search(header_txt.replace(" ", ""))  # header has "OPEN TO CLOSE DATE:12/06/2025 -01/05/2026"
        if not m:
            # pdfplumber sometimes leaves extra spacing; try a looser match
            m = re.search(r"(\d{2}/\d{2}/\d{4}).{0,5}-.{0,5}(\d{2}/\d{2}/\d{4})", header_txt)
        if not m:
            raise ValueError(f"Could not find statement window in {pdf_path.name}")
        start = pd.to_datetime(m.group(1))
        end   = pd.to_datetime(m.group(2))

        # Transactions live on page 3 only for these statements,
        # but we iterate every page defensively.
        in_purchases = False
        for page in pdf.pages:
            text = page.extract_text() or ""
            for line in text.splitlines():
                # Section boundaries
                if "PURCHASES" in line.upper() and "MERCHANTCATEGORY" in line.upper().replace(" ", ""):
                    in_purchases = True
                    continue
                if "FeesandInterestCharged" in line.replace(" ", ""):
                    in_purchases = False
                    continue
                if not in_purchases:
                    continue

                match = DISCOVER_TXN_RX.match(line)
                if not match:
                    continue
                mm_dd, body, amt = match.groups()
                body = body.strip()
                # The last token group in `body` should be one of the known
                # Discover categories; splitting on it separates merchant
                # description from the category even when the merchant name
                # contains phone numbers or other noisy tokens.
                cat = None
                desc = None
                for known in DISCOVER_KNOWN_CATS:
                    if body.endswith(" " + known) or body == known:
                        cat = known
                        desc = body[: len(body) - len(known)].strip()
                        break
                if cat is None or not desc:
                    continue

                year = _infer_year(mm_dd, start, end)
                date = pd.Timestamp(year=year, month=int(mm_dd.split("/")[0]),
                                    day=int(mm_dd.split("/")[1]))
                rows.append({
                    "Date":           date.strftime("%Y-%m-%d"),
                    "Description":   _clean_description(desc),
                    "Category":      cat,
                    "Amount":        float(amt.replace(",", "")),
                    "Payment Method": "Discover Card",
                    "Source":        pdf_path.name,
                })
    return rows


# ------------------------------------------------------------------
# Parser: BofA Visa Signature credit card
# ------------------------------------------------------------------
def parse_bofa_visa(pdf_path: Path) -> list[dict]:
    rows: list[dict] = []
    with pdfplumber.open(pdf_path) as pdf:
        # Period: pull from page 1's "December 28 - January 27, 2026" line
        p1 = pdf.pages[0].extract_text() or ""
        start, end = _parse_bofa_period(p1)

        in_purchases = False
        for page in pdf.pages:
            text = page.extract_text() or ""
            for line in text.splitlines():
                # Section boundaries
                low = line.lower()
                if "purchases and adjustments" in low and "total" not in low:
                    in_purchases = True
                    continue
                if low.startswith("total purchases and adjustments"):
                    in_purchases = False
                    continue
                if "interest charged" in low and not in_purchases:
                    continue
                if not in_purchases:
                    continue

                match = BOFA_TXN_RX.match(line)
                if not match:
                    continue
                trans_mm_dd, _post_mm_dd, desc, _ref, _acct, amt = match.groups()
                amount = float(amt.replace(",", ""))
                if amount <= 0:
                    # Adjustments/returns would be negative — skip for expense view
                    continue

                year = _infer_year(trans_mm_dd, start, end)
                date = pd.Timestamp(year=year, month=int(trans_mm_dd.split("/")[0]),
                                    day=int(trans_mm_dd.split("/")[1]))
                rows.append({
                    "Date":           date.strftime("%Y-%m-%d"),
                    "Description":   _clean_description(desc),
                    "Category":      "Uncategorized",
                    "Amount":        amount,
                    "Payment Method": "BofA Credit",
                    "Source":        pdf_path.name,
                })
    return rows


def _parse_bofa_period(page_text: str) -> tuple[pd.Timestamp, pd.Timestamp]:
    """Extract the statement period from a BofA credit-card page 1.
    Supports both 'December 28 - January 27, 2026' and 'January 28 - February 27, 2026'.
    """
    MONTHS = {m: i for i, m in enumerate(
        ["January","February","March","April","May","June","July",
         "August","September","October","November","December"], start=1)}
    for line in page_text.splitlines():
        m = BOFA_PERIOD_RX.search(line)
        if not m:
            continue
        mon_a, day_a, mon_b, day_b, year = m.groups()
        mon_b = mon_b or mon_a      # "January 28 - 27, 2026" style isn't expected, but be safe
        end_year = int(year)
        # If end is January and start is December, the start belongs to the previous year
        start_year = end_year - 1 if (MONTHS[mon_b] == 1 and MONTHS[mon_a] == 12) else end_year
        start = pd.Timestamp(year=start_year, month=MONTHS[mon_a], day=int(day_a))
        end   = pd.Timestamp(year=end_year,   month=MONTHS[mon_b], day=int(day_b))
        return start, end
    raise ValueError("Could not find BofA statement period in page 1")


# ------------------------------------------------------------------
# Detection + description cleanup
# ------------------------------------------------------------------
def detect_kind(pdf_path: Path) -> str:
    """Return one of: 'discover', 'bofa_visa', 'bofa_checking', 'unknown'."""
    name = pdf_path.name
    # Discover statements are named after the month
    if re.match(r"^(January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{4}\.pdf$", name):
        return "discover"
    # BofA eStmt_YYYY-MM-DD.pdf — use page 1 to disambiguate Visa vs checking
    if name.startswith("eStmt_"):
        with pdfplumber.open(pdf_path) as pdf:
            p1 = pdf.pages[0].extract_text() or ""
        if "Visa Signature" in p1 or "Purchases and Adjustments" in p1:
            return "bofa_visa"
        if "SafeBalance" in p1 or "ATM and debit card" in p1:
            return "bofa_checking"
    return "unknown"


def _clean_description(desc: str) -> str:
    """Trim noisy tokens that BofA/Discover include in raw descriptions."""
    desc = desc.strip()
    # Collapse runs of whitespace
    desc = re.sub(r"\s+", " ", desc)
    # Drop trailing phone numbers (800-xxx-xxxx, 866-xxx-xxxx, etc.)
    desc = re.sub(r"\s+\d{3}-\d{3}-\d{4}(\s+[A-Z]{2})?\s*$", "", desc)
    return desc


# ------------------------------------------------------------------
# Main
# ------------------------------------------------------------------
def main():
    all_rows: list[dict] = []
    skipped: list[str] = []

    pdfs = sorted(WORKSPACE.glob("*.pdf"))
    print(f"Found {len(pdfs)} PDFs under {WORKSPACE}\n")

    for pdf_path in pdfs:
        kind = detect_kind(pdf_path)
        if kind == "discover":
            rows = parse_discover(pdf_path)
            print(f"  Discover      {pdf_path.name:<30} -> {len(rows):>3} txns  "
                  f"${sum(r['Amount'] for r in rows):>9,.2f}")
            all_rows.extend(rows)
        elif kind == "bofa_visa":
            rows = parse_bofa_visa(pdf_path)
            print(f"  BofA Visa     {pdf_path.name:<30} -> {len(rows):>3} txns  "
                  f"${sum(r['Amount'] for r in rows):>9,.2f}")
            all_rows.extend(rows)
        elif kind == "bofa_checking":
            print(f"  BofA Checking {pdf_path.name:<30} -> skipped (stub: layout differs, left for next iteration)")
            skipped.append(pdf_path.name)
        else:
            print(f"  UNKNOWN       {pdf_path.name:<30} -> skipped")
            skipped.append(pdf_path.name)

    df = pd.DataFrame(all_rows).sort_values("Date").reset_index(drop=True)
    df.to_csv(OUT_CSV, index=False)

    print(f"\nWrote {len(df)} rows -> {OUT_CSV}")
    if len(df):
        print(f"Date range: {df['Date'].min()} -> {df['Date'].max()}")
        print(f"Total extracted spend: ${df['Amount'].sum():,.2f}")
        print("\nBy payment method:")
        print(df.groupby("Payment Method")["Amount"]
                .agg(["sum", "count"]).round(2))
        print("\nDiscover-only category mix (BofA rows are Uncategorized):")
        disc = df[df["Payment Method"] == "Discover Card"]
        if len(disc):
            print(disc.groupby("Category")["Amount"]
                    .agg(["sum", "count"]).sort_values("sum", ascending=False).round(2))
    if skipped:
        print(f"\nSkipped {len(skipped)} PDFs (checking-account or unknown layout):")
        for s in skipped:
            print(f"  · {s}")


if __name__ == "__main__":
    main()
