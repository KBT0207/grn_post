import pandas as pd

from config import COLUMNS

REQUIRED_HEADERS = list(COLUMNS.values())

_TEXT_FIELDS = (
    "voucher_type", "voucher_no", "date", "reference_no", "party_name",
    "purchase_ledger", "item_name", "unit", "tracking_no", "godown", "narration",
)
_NUMERIC_FIELDS = ("qty", "rate", "amount")


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

    df = df[REQUIRED_HEADERS].copy()
    df = df.rename(columns={header: field for field, header in COLUMNS.items()})

    df = df.dropna(how="all").reset_index(drop=True)

    for field in _TEXT_FIELDS:
        df[field] = df[field].fillna("").astype(str).str.strip()

    for field in _NUMERIC_FIELDS:
        df[field] = pd.to_numeric(df[field], errors="coerce")

    errors = []

    bad_numeric = df[df["qty"].isna() | df["rate"].isna() | df["amount"].isna()]
    if not bad_numeric.empty:
        errors.append(f"Non-numeric Qty/Rate/Amount in Excel row(s): {[i + 2 for i in bad_numeric.index]}")

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

        vouchers.append({
            "voucher_type": voucher_type,
            "voucher_no": voucher_no,
            "date": date_str,
            "reference_no": reference_no,
            "party_name": party_name,
            "tracking_no": tracking_nos[0],
            "narration": narration,
            "items": group.to_dict("records"),
        })

    return vouchers
