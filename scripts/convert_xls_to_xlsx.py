"""指定フォルダ内の .xls を Excel COM で .xlsx に一括変換する（書式維持）。

実行例:
  python scripts/convert_xls_to_xlsx.py --src "C:\\Users\\USER\\Documents\\三浦さんマージン清算\\カルチャーキッズマージン明細" --keep-original
  python scripts/convert_xls_to_xlsx.py --src "..." --delete-original  # 元の.xlsを削除
"""
from __future__ import annotations
import argparse
import os
import sys


def convert_one(excel, src_xls: str, dst_xlsx: str) -> None:
    """1ファイル変換（FileFormat=51 = xlOpenXMLWorkbook）"""
    wb = excel.Workbooks.Open(os.path.abspath(src_xls))
    try:
        wb.SaveAs(os.path.abspath(dst_xlsx), FileFormat=51)
    finally:
        wb.Close(SaveChanges=False)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--src", required=True, help="変換元フォルダ")
    ap.add_argument("--dst", help="変換先フォルダ（省略時は src と同じ）")
    ap.add_argument("--keep-original", action="store_true", help="元の.xlsを残す（デフォルト）")
    ap.add_argument("--delete-original", action="store_true", help="変換後に元の.xlsを削除")
    args = ap.parse_args()

    src = args.src
    dst = args.dst or src
    os.makedirs(dst, exist_ok=True)

    targets = []
    for fn in os.listdir(src):
        if fn.endswith(".xls") and not fn.endswith(".xlsx"):
            if "_bak_" in fn:
                continue
            targets.append(fn)
    targets.sort()
    print(f"=== {len(targets)} files to convert ===")

    if not targets:
        return

    import pythoncom  # type: ignore
    import win32com.client  # type: ignore

    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        for i, fn in enumerate(targets, 1):
            src_p = os.path.join(src, fn)
            dst_p = os.path.join(dst, fn[:-4] + ".xlsx")
            print(f"  [{i}/{len(targets)}] {fn}")
            convert_one(excel, src_p, dst_p)
            if args.delete_original and os.path.exists(dst_p):
                os.remove(src_p)
                print(f"      removed: {fn}")
    finally:
        excel.Quit()
        pythoncom.CoUninitialize()
    print("DONE")


if __name__ == "__main__":
    main()
