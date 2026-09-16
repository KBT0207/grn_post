import pandas as pd
import re
from datetime import datetime

from config import DATE_INPUT_FORMAT

_XML_ESCAPES = {"&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&apos;"}
_ESCAPE_RE = re.compile("|".join(re.escape(c) for c in _XML_ESCAPES))


def xml_escape(value) -> str:
    text = "" if value is None else str(value)
    return _ESCAPE_RE.sub(lambda m: _XML_ESCAPES[m.group(0)], text)



def to_tally_date(date_str, input_format="%d-%m-%Y"):
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


def build_voucher(voucher: dict) -> str:
    date_tally = to_tally_date(voucher["date"])
    entries = "\n".join(build_inventory_entry(item) for item in voucher["items"])

    # Receipt Note is a plain inventory voucher (not Invoice mode), and this
    # company's GRNs are routinely posted with no party attached yet - so
    # only emit the party tags when a Party A/c Name was actually given.
    party_block = ""
    if voucher["party_name"]:
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
{entries}
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
