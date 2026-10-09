import pandas as pd

from config import COLUMNS, OPTIONAL_FIELDS

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
