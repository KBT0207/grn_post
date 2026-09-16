# GRN → Tally Receipt Note Importer

Reads a Receipt Note (GRN) Excel file and posts it straight into whichever
Tally company is currently open, as `Receipt Note` vouchers, via Tally's
HTTP-XML server. No lookups against Tally — it trusts that every ledger,
stock item, unit and godown in the Excel already exists.

## Setup

```bash
pip install -r requirements.txt
```

In Tally, make sure the XML server is on (**F1 > Settings > Connectivity**,
or check your port in **F12 configuration**). Default assumed here is
`http://localhost:9000`.

## Excel format expected

One row per item line. Every row belonging to the same voucher must repeat
identical header values (Voucher Type, Receipt No, Date, Reference No,
Party A/c Name) and the same Tracking No. **Party A/c Name may be left
blank** — matches how this company actually posts GRNs (party gets attached
later at the Purchase Invoice stage); when blank, the script simply omits
`PARTYNAME`/`PARTYLEDGERNAME` from the XML instead of sending empty tags.

| Voucher Type | Receipt No | Date | Reference No | Party A/c Name | Purchase Ledger | Item Name | Unit | Tracking No | Qty | Rate | Amount | Godown | Narration |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| Receipt Note | 1 | 01-04-2026 | 1 | | Purchase Fruits & Veg | Red Jaam | Kg | 1 | 25 | 100 | 2500 | Main Location | |
| Receipt Note | 1 | 01-04-2026 | 1 | | Purchase Fruits & Veg | White Jaam | Kg | 1 | 20 | 100 | 2000 | Main Location | |

Date format is `DD-MM-YYYY` (change `DATE_INPUT_FORMAT` in `config.py` if
yours differs).

## Usage

Always look at the XML before you trust it with real data:

```bash
python main.py your_file.xlsx --dry-run
```

Import for real (posts one combined request with all vouchers):

```bash
python main.py your_file.xlsx
```

Import one voucher per HTTP request instead — slower, but one bad voucher
won't block the rest, and you get a per-voucher pass/fail line:

```bash
python main.py your_file.xlsx --one-voucher-per-request
```

Other flags:
- `--url http://localhost:9000` — point at a different Tally XML port
- `--company "Kay Bee Exports International Pvt Ltd"` — target a specific
  company instead of whatever's currently open in Tally
- `--sheet "Sheet2"` — pick a non-default sheet

## What the script validates (Excel-side only, nothing checked in Tally)

- Required columns are present
- Qty / Rate / Amount are numeric
- No blank Receipt No, Date, Purchase Ledger, Item Name, Tracking No, or
  Godown (Party A/c Name is allowed to be blank)
- All item lines of one voucher share the same Tracking No.

Anything it doesn't validate (ledger names, item names, units, godowns
actually existing, GST setup, etc.) is passed straight through to Tally —
by design, since you said everything's already set up there. If a name is
wrong, Tally will reject that voucher and the response's `LINEERROR` text
(printed by `main.py`) will say so.

## XML shape (per voucher)

```xml
<VOUCHER VCHTYPE="Receipt Note" ACTION="Create">
  <DATE>20260401</DATE>
  <EFFECTIVEDATE>20260401</EFFECTIVEDATE>
  <VOUCHERTYPENAME>Receipt Note</VOUCHERTYPENAME>
  <VOUCHERNUMBER>1</VOUCHERNUMBER>
  <REFERENCE>1</REFERENCE>
  <!-- PARTYLEDGERNAME / PARTYNAME only appear when Party A/c Name is non-blank -->
  <NARRATION></NARRATION>
  <ALLINVENTORYENTRIES.LIST>
    <STOCKITEMNAME>Red Jaam</STOCKITEMNAME>
    <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>
    <RATE>100.00/Kg</RATE>
    <AMOUNT>-2500.00</AMOUNT>
    <ACTUALQTY>25 Kg</ACTUALQTY>
    <BILLEDQTY>25 Kg</BILLEDQTY>
    <BATCHALLOCATIONS.LIST>
      <GODOWNNAME>Main Location</GODOWNNAME>
      <BATCHNAME>Primary Batch</BATCHNAME>
      <TRACKINGNUMBER>1</TRACKINGNUMBER>
      <AMOUNT>-2500.00</AMOUNT>
      <ACTUALQTY>25 Kg</ACTUALQTY>
      <BILLEDQTY>25 Kg</BILLEDQTY>
    </BATCHALLOCATIONS.LIST>
    <ACCOUNTINGALLOCATIONS.LIST>
      <LEDGERNAME>Purchase Fruits &amp; Veg</LEDGERNAME>
      <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>
      <AMOUNT>-2500.00</AMOUNT>
    </ACCOUNTINGALLOCATIONS.LIST>
  </ALLINVENTORYENTRIES.LIST>
</VOUCHER>
```

## Please test one voucher before a bulk run

Your `TRACKINGNUMBER`-on-`BATCHALLOCATIONS.LIST` field comes from a
company-specific TDL, and TDL customizations sometimes change exactly which
tags a voucher screen expects underneath (e.g. whether `ACCOUNTINGALLOCATIONS.LIST`
per item is actually needed, or whether `OBJVIEW`/`ISINVOICE` should differ).
Everything above matches your screenshot and column layout, but the fastest
way to confirm it's 100% right for *your* TDL is:

1. Make a 1-voucher, 1-item test Excel file
2. `python main.py test.xlsx --dry-run` and eyeball the XML
3. `python main.py test.xlsx` for real, then open that voucher in Tally and
   check the Tracking No., Godown, Qty, Rate, and Purchase Ledger all landed
   correctly
4. Only then run your full file

If Tally rejects it, the printed `LINEERROR` text tells you exactly which
tag/value it didn't like — that's the fastest way to pin down any TDL
differences from what's above.
