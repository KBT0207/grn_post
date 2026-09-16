from config import TALLY_URL
from excel_reader import read_grn_excel, group_vouchers
from xml_builder import build_import_envelope
from tally_client import post_xml, parse_import_response


def main(excel_path: str, company: str = "", url: str = TALLY_URL):
    df = read_grn_excel(excel_path, sheet_name=0)
    vouchers = group_vouchers(df)

    if not vouchers:
        print("No vouchers found in the Excel file.")
        return

    total_items = len(df)
    total_amount = sum(item["amount"] for v in vouchers for item in v["items"])

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
        voucher_amount = sum(item["amount"] for item in v["items"])

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
    main("sample_grn1.xlsx")