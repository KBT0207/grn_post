"""
GRN Excel -> Tally (XML) Importer  (single-file version, with Tax Ledger support)
=================================================================================
Run as:

    python grn_import.py path/to/sample_grn_with_tax.xlsx
    python grn_import.py path/to/file.xlsx --company "My Company" --url http://localhost:9000
"""

import os
import re
import sys
import argparse
from datetime import datetime

import pandas as pd
import requests
import xml.etree.ElementTree as ET


# ============================================================================
# SECTION 1: CONFIG
# ============================================================================

# Tally's XML server (Gateway of Tally > F1 > Settings > Connectivity, or F12
# config, shows the port; 9000 is TallyPrime's default). Override with the
# TALLY_URL env var, or pass --url on the command line.
TALLY_URL = os.getenv("TALLY_URL", "http://localhost:9000")

# The only voucher type this project imports.
VOUCHER_TYPE = "Receipt Note"

# Excel column headers -> internal field names. Edit the right-hand side if
# your sheet's headers ever change; nothing else in the code needs to change.
COLUMNS = {
    "voucher_type": "Voucher Type",
    "voucher_no": "Receipt No",
    "date": "Date",
    "reference_no": "Reference No",
    "party_name": "Party A/c Name",
    "purchase_ledger": "Purchase Ledger",
    "item_name": "Item Name",
    "unit": "Unit",
    "tracking_no": "Tracking No",
    "qty": "Qty",
    "rate": "Rate",
    "amount": "Amount",
    "godown": "Godown",
    "narration": "Narration",
    "tax_ledger": "Tax Ledger",     # NEW (optional column)
    "tax_amount": "Tax Amount",     # NEW (optional column)
    "total_amount": "Total Amount", # NEW (optional column) voucher total = items + tax
}

# Columns that may be missing from the sheet entirely (older sheets still work).
OPTIONAL_FIELDS = ("tax_ledger", "tax_amount", "total_amount")

# XML tag used for the tax-ledger and party-total lines inside the voucher. If Tally ignores the
# tax line, try "ALLLEDGERENTRIES.LIST" here instead.
TAX_ENTRY_TAG = "LEDGERENTRIES.LIST"

# How dates are written in the source Excel, e.g. "01-04-2026"
DATE_INPUT_FORMAT = "%d-%m-%Y"
# DATE_INPUT_FORMAT = "%Y-%m-%d"


# ============================================================================
# SECTION 2: EXCEL READING
# ============================================================================

REQUIRED_HEADERS = [h for f, h in COLUMNS.items() if f not in OPTIONAL_FIELDS]

_TEXT_FIELDS = (
    "voucher_type", "voucher_no", "date", "reference_no", "party_name",
    "purchase_ledger", "item_name", "unit", "tracking_no", "godown", "narration",
    "tax_ledger",
)
_NUMERIC_FIELDS = ("qty", "rate", "amount", "tax_amount", "total_amount")


def read_grn_excel(path: str, sheet_name=0) -> pd.DataFrame:
    """Load the GRN sheet, rename headers to internal field names, and
    validate that every row has what it needs to build a voucher."""
    df = pd.read_excel(path, sheet_name=sheet_name, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]

    missing = [h for h in REQUIRED_HEADERS if h not in df.columns]
    if missing:
        raise ValueError(
            f"Excel is missing required column(s): {missing}. "
            f"Found columns: {list(df.columns)}"
        )

    # Optional columns (tax): create empty ones if the sheet doesn't have them.
    for f in OPTIONAL_FIELDS:
        if COLUMNS[f] not in df.columns:
            df[COLUMNS[f]] = None

    df = df[list(COLUMNS.values())].copy()
    df = df.rename(columns={header: field for field, header in COLUMNS.items()})

    df = df.dropna(how="all").reset_index(drop=True)

    for field in _TEXT_FIELDS:
        df[field] = df[field].fillna("").astype(str).str.strip()

    for field in _NUMERIC_FIELDS:
        df[field] = pd.to_numeric(df[field], errors="coerce")

    errors = []

    # Qty must always be a real number - Tally can't post a line with no quantity.
    bad_qty = df[df["qty"].isna()]
    if not bad_qty.empty:
        errors.append(f"Non-numeric/blank Qty in Excel row(s): {[i + 2 for i in bad_qty.index]}")

    # Rate/Amount/Tax Amount are allowed to be blank or zero (e.g. free-of-cost
    # items, samples, no-tax rows) - blank cells default to 0.
    df["rate"] = df["rate"].fillna(0)
    df["amount"] = df["amount"].fillna(0)
    df["tax_amount"] = df["tax_amount"].fillna(0)
    df["total_amount"] = df["total_amount"].fillna(0)

    # A tax amount with no ledger name can't be posted.
    bad_tax = df[(df["tax_amount"] != 0) & (df["tax_ledger"] == "")]
    if not bad_tax.empty:
        errors.append(
            f"Tax Amount given but Tax Ledger blank in Excel row(s): {[i + 2 for i in bad_tax.index]}"
        )

    for field, label in (
        ("voucher_no", "Receipt No"),
        ("date", "Date"),
        ("purchase_ledger", "Purchase Ledger"),
        ("item_name", "Item Name"),
        ("tracking_no", "Tracking No"),
        ("godown", "Godown"),
    ):
        blanks = df[df[field] == ""]
        if not blanks.empty:
            errors.append(f"Blank '{label}' in Excel row(s): {[i + 2 for i in blanks.index]}")

    if errors:
        raise ValueError("Excel validation failed:\n  - " + "\n  - ".join(errors))

    return df


def group_vouchers(df: pd.DataFrame) -> list[dict]:
    key_cols = ["voucher_type", "voucher_no", "date", "reference_no", "party_name"]
    vouchers = []

    for key, group in df.groupby(key_cols, sort=False):
        voucher_type, voucher_no, date_str, reference_no, party_name = key

        tracking_nos = group["tracking_no"].unique().tolist()
        if len(tracking_nos) > 1:
            raise ValueError(
                f"Voucher {voucher_no!r} (Date {date_str}) has mixed Tracking No "
                f"values {tracking_nos}; every item in a voucher must share the "
                f"same Tracking No."
            )

        narration = next((n for n in group["narration"] if n), "")

        # Tax ledgers for this voucher: {ledger name: total amount}.
        # Enter the tax once per voucher; repeated rows of the same ledger add up.
        taxes = {}
        for _, r in group.iterrows():
            if r["tax_ledger"] and r["tax_amount"]:
                taxes[r["tax_ledger"]] = taxes.get(r["tax_ledger"], 0) + float(r["tax_amount"])

        # Voucher total = items + tax. The Excel "Total Amount" (entered once per
        # voucher) must agree with it; if left blank the calculated value is used.
        items_total = float(group["amount"].sum())
        calc_total = round(items_total + sum(taxes.values()), 2)
        excel_total = round(float(group["total_amount"].sum()), 2)
        if excel_total and abs(excel_total - calc_total) > 0.01:
            raise ValueError(
                f"Voucher {voucher_no!r} (Date {date_str}): Total Amount in Excel is "
                f"{excel_total:.2f} but items + tax = {calc_total:.2f} "
                f"(items {items_total:.2f} + tax {sum(taxes.values()):.2f})."
            )
        total = excel_total or calc_total

        vouchers.append({
            "voucher_type": voucher_type,
            "voucher_no": voucher_no,
            "date": date_str,
            "reference_no": reference_no,
            "party_name": party_name,
            "tracking_no": tracking_nos[0],
            "narration": narration,
            "items": group.to_dict("records"),
            "taxes": taxes,
            "total": total,
        })

    return vouchers


# ============================================================================
# SECTION 3: XML BUILDING
# ============================================================================

_XML_ESCAPES = {"&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&apos;"}
_ESCAPE_RE = re.compile("|".join(re.escape(c) for c in _XML_ESCAPES))


def xml_escape(value) -> str:
    text = "" if value is None else str(value)
    return _ESCAPE_RE.sub(lambda m: _XML_ESCAPES[m.group(0)], text)



def to_tally_date(date_str, input_format=DATE_INPUT_FORMAT):
    # Case 1: already a real datetime/Timestamp (typical when Excel col is date-formatted)
    if isinstance(date_str, (datetime, pd.Timestamp)):
        return date_str.strftime("%Y%m%d")

    # Case 2: it's a string — strip time portion if present, then try formats
    s = str(date_str).strip()

    # If it looks like '2026-09-16 00:00:00' or '2026-09-16', handle ISO directly
    for fmt in (input_format, "%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(s, fmt).strftime("%Y%m%d")
        except ValueError:
            continue

    raise ValueError(f"Unrecognized date format: {date_str!r}")


def _fmt_qty(qty: float, unit: str) -> str:
    # Tally wants e.g. "25 Kg" / "25.5 Kg" - no trailing .0 for whole numbers
    qty_str = f"{qty:g}"
    return f"{qty_str} {unit}".strip()


def _fmt_rate(rate: float, unit: str) -> str:
    return f"{rate:.2f}/{unit}".strip()


def _fmt_amount(amount: float) -> str:
    return f"{amount:.2f}"


def build_inventory_entry(item: dict) -> str:
    qty_str = _fmt_qty(item["qty"], item["unit"])
    rate_str = _fmt_rate(item["rate"], item["unit"])
    # Inward stock movement -> ISDEEMEDPOSITIVE=Yes with a negative AMOUNT,
    # per Tally's standard debit/credit convention for inventory entries.
    neg_amount = _fmt_amount(-float(item["amount"]))

    return f"""        <ALLINVENTORYENTRIES.LIST>
            <STOCKITEMNAME>{xml_escape(item['item_name'])}</STOCKITEMNAME>
            <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>
            <RATE>{xml_escape(rate_str)}</RATE>
            <AMOUNT>{neg_amount}</AMOUNT>
            <ACTUALQTY>{xml_escape(qty_str)}</ACTUALQTY>
            <BILLEDQTY>{xml_escape(qty_str)}</BILLEDQTY>
            <BATCHALLOCATIONS.LIST>
                <GODOWNNAME>{xml_escape(item['godown'])}</GODOWNNAME>
                <BATCHNAME>Primary Batch</BATCHNAME>
                <TRACKINGNUMBER>{xml_escape(item['tracking_no'])}</TRACKINGNUMBER>
                <AMOUNT>{neg_amount}</AMOUNT>
                <ACTUALQTY>{xml_escape(qty_str)}</ACTUALQTY>
                <BILLEDQTY>{xml_escape(qty_str)}</BILLEDQTY>
            </BATCHALLOCATIONS.LIST>
            <ACCOUNTINGALLOCATIONS.LIST>
                <LEDGERNAME>{xml_escape(item['purchase_ledger'])}</LEDGERNAME>
                <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>
                <AMOUNT>{neg_amount}</AMOUNT>
            </ACCOUNTINGALLOCATIONS.LIST>
        </ALLINVENTORYENTRIES.LIST>"""


def build_tax_entry(ledger: str, amount: float) -> str:
    # Same sign convention as the item lines: ISDEEMEDPOSITIVE=Yes, negative AMOUNT.
    # If Tally shows the wrong sign, flip it here.
    neg_amount = _fmt_amount(-float(amount))

    return f"""        <{TAX_ENTRY_TAG}>
            <LEDGERNAME>{xml_escape(ledger)}</LEDGERNAME>
            <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>
            <LEDGERFROMITEM>No</LEDGERFROMITEM>
            <REMOVEZEROENTRIES>No</REMOVEZEROENTRIES>
            <ISPARTYLEDGER>No</ISPARTYLEDGER>
            <ISLASTDEEMEDPOSITIVE>Yes</ISLASTDEEMEDPOSITIVE>
            <AMOUNT>{neg_amount}</AMOUNT>
        </{TAX_ENTRY_TAG}>"""


def build_party_entry(party: str, total: float) -> str:
    # Party (supplier) ledger line carrying the voucher TOTAL (items + tax).
    # Opposite sign to the item/tax lines, so the voucher balances.
    return f"""        <{TAX_ENTRY_TAG}>
            <LEDGERNAME>{xml_escape(party)}</LEDGERNAME>
            <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
            <LEDGERFROMITEM>No</LEDGERFROMITEM>
            <REMOVEZEROENTRIES>No</REMOVEZEROENTRIES>
            <ISPARTYLEDGER>Yes</ISPARTYLEDGER>
            <ISLASTDEEMEDPOSITIVE>No</ISLASTDEEMEDPOSITIVE>
            <AMOUNT>{_fmt_amount(float(total))}</AMOUNT>
        </{TAX_ENTRY_TAG}>"""


def build_voucher(voucher: dict) -> str:
    date_tally = to_tally_date(voucher["date"])
    entries = "\n".join(build_inventory_entry(item) for item in voucher["items"])
    # Tax ledger lines, posted right after the item lines.
    tax_entries = "\n".join(
        build_tax_entry(name, amt) for name, amt in voucher.get("taxes", {}).items()
    )

    # Receipt Note is a plain inventory voucher (not Invoice mode), and this
    # company's GRNs are routinely posted with no party attached yet - so
    # only emit the party tags when a Party A/c Name was actually given.
    party_block = ""
    party_total_entry = ""
    if voucher["party_name"]:
        party_total_entry = build_party_entry(voucher["party_name"], voucher.get("total", 0)) + "\n"
        party_block = (
            f"        <PARTYLEDGERNAME>{xml_escape(voucher['party_name'])}</PARTYLEDGERNAME>\n"
            f"        <PARTYNAME>{xml_escape(voucher['party_name'])}</PARTYNAME>\n"
        )

    return f"""    <TALLYMESSAGE xmlns:UDF="TallyUDF">
      <VOUCHER VCHTYPE="{xml_escape(voucher['voucher_type'])}" ACTION="Create">
        <DATE>{date_tally}</DATE>
        <EFFECTIVEDATE>{date_tally}</EFFECTIVEDATE>
        <VOUCHERTYPENAME>{xml_escape(voucher['voucher_type'])}</VOUCHERTYPENAME>
        <VOUCHERNUMBER>{xml_escape(voucher['voucher_no'])}</VOUCHERNUMBER>
        <REFERENCE>{xml_escape(voucher['reference_no'])}</REFERENCE>
{party_block}        <NARRATION>{xml_escape(voucher['narration'])}</NARRATION>
{party_total_entry}{entries}
{tax_entries}
      </VOUCHER>
    </TALLYMESSAGE>"""


def build_import_envelope(vouchers: list[dict], company_name: str = "") -> str:
    company_block = (
        f"<SVCURRENTCOMPANY>{xml_escape(company_name)}</SVCURRENTCOMPANY>"
        if company_name else ""
    )
    messages = "\n".join(build_voucher(v) for v in vouchers)

    return f"""<ENVELOPE>
  <HEADER>
    <TALLYREQUEST>Import Data</TALLYREQUEST>
  </HEADER>
  <BODY>
    <IMPORTDATA>
      <REQUESTDESC>
        <REPORTNAME>Vouchers</REPORTNAME>
        <STATICVARIABLES>
          {company_block}
        </STATICVARIABLES>
      </REQUESTDESC>
      <REQUESTDATA>
{messages}
      </REQUESTDATA>
    </IMPORTDATA>
  </BODY>
</ENVELOPE>"""


# ============================================================================
# SECTION 4: TALLY CLIENT
# ============================================================================

def post_xml(xml_str: str, url: str = TALLY_URL, timeout: int = 120) -> str:
    resp = requests.post(
        url,
        data=xml_str.encode("utf-8"),
        headers={"Content-Type": "text/xml; charset=utf-8"},
        timeout=timeout,
    )
    resp.raise_for_status()
    return resp.text


def parse_import_response(response_text: str) -> dict:
    try:
        root = ET.fromstring(response_text)
    except ET.ParseError:
        return {
            "created": None, "altered": None, "errors": None,
            "exceptions": None, "line_errors": [], "raw": response_text,
        }

    def _int(tag):
        text = root.findtext(f".//{tag}")
        return int(text.strip()) if text and text.strip().lstrip("-").isdigit() else None

    line_errors = [e.text.strip() for e in root.iter("LINEERROR") if e.text and e.text.strip()]

    return {
        "created": _int("CREATED"),
        "altered": _int("ALTERED"),
        "errors": _int("ERRORS"),
        "exceptions": _int("EXCEPTIONS"),
        "line_errors": line_errors,
        "raw": response_text,
    }


# ============================================================================
# SECTION 5: MAIN
# ============================================================================

def main(excel_path: str, company: str = "", url: str = TALLY_URL):
    df = read_grn_excel(excel_path, sheet_name=0)
    vouchers = group_vouchers(df)

    if not vouchers:
        print("No vouchers found in the Excel file.")
        return

    total_items = len(df)
    total_amount = sum(v["total"] for v in vouchers)

    print("=" * 60)
    print(f"GRN Import Started")
    print(f"  Excel file   : {excel_path}")
    print(f"  Tally URL    : {url}")
    print(f"  Company      : {company or '(currently open company)'}")
    print(f"  Vouchers     : {len(vouchers)}")
    print(f"  Item lines   : {total_items}")
    print(f"  Total amount : {total_amount:.2f}")
    print("=" * 60)

    ok, failed = 0, []

    for idx, v in enumerate(vouchers, start=1):
        item_count = len(v["items"])
        voucher_amount = v["total"]

        print(f"\n[{idx}/{len(vouchers)}] Voucher {v['voucher_no']}  "
              f"(Date: {v['date']}, Party: {v['party_name'] or '-'}, "
              f"Items: {item_count}, Amount: {voucher_amount:.2f})")

        xml_str = build_import_envelope([v], company_name=company)

        try:
            result = parse_import_response(post_xml(xml_str, url=url))
        except Exception as exc:
            failed.append(v["voucher_no"])
            print(f"    -> FAILED: could not reach Tally ({exc})")
            continue

        if result["errors"] or result["line_errors"]:
            failed.append(v["voucher_no"])
            print(f"    -> FAILED")
            for err in (result["line_errors"] or [result["raw"]]):
                print(f"       - {err}")
        else:
            ok += 1
            print(f"    -> OK  (Created: {result['created']}, Altered: {result['altered']})")

    print("\n" + "=" * 60)
    print(f"Import Finished: {ok} succeeded, {len(failed)} failed out of {len(vouchers)}")
    if failed:
        print(f"Failed voucher number(s): {', '.join(failed)}")
    print("=" * 60)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Import GRN Excel rows into Tally as Receipt Note vouchers.")
    parser.add_argument("excel_path", nargs="?", default="sample_grn1.xlsx", help="Path to the GRN Excel file")
    parser.add_argument("--company", default="", help="Tally company name (blank = currently open company)")
    parser.add_argument("--url", default=TALLY_URL, help="Tally XML server URL (default: %(default)s)")
    args = parser.parse_args()

    main(args.excel_path, company=args.company, url=args.url)
