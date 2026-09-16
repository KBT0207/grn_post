import os

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
}

# How dates are written in the source Excel, e.g. "01-04-2026"
DATE_INPUT_FORMAT = "%d-%m-%Y"
# DATE_INPUT_FORMAT = "%Y-%m-%d"


