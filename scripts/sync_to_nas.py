"""ローカルの xlsx を NAS の カルチャーキッズマージン明細 にアップロードする。

NAS側の既存 .xls は同フォルダの _xls_backup_YYYYMMDD/ に退避してから、
ローカルの .xlsx で上書き／追加する。
"""
from __future__ import annotations
import argparse
import os
import shutil
from datetime import datetime


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--src", default=r"C:\Users\USER\Documents\三浦さんマージン清算\カルチャーキッズマージン明細")
    ap.add_argument("--dst", default=r"Y:\_★20170701作業用\9三浦\代理店マージン精算書(2021.6～）\カルチャーキッズマージン明細")
    ap.add_argument("--apply", action="store_true")
    args = ap.parse_args()

    src, dst = args.src, args.dst
    if not os.path.isdir(src):
        print(f"NG: src not found: {src}")
        return
    if not os.path.isdir(dst):
        print(f"NG: dst not found: {dst}")
        return

    # 1) NAS側 .xls を _xls_backup_YYYYMMDD/ に退避
    today = datetime.now().strftime("%Y%m%d")
    bak_dir = os.path.join(dst, f"_xls_backup_{today}")
    nas_xls = [f for f in os.listdir(dst) if f.endswith(".xls") and "精算書" in f]
    print(f"=== STEP 1: NAS側 .xls を退避 ({len(nas_xls)}件) → {bak_dir} ===")
    if args.apply:
        os.makedirs(bak_dir, exist_ok=True)
    for fn in nas_xls:
        src_p = os.path.join(dst, fn)
        bak_p = os.path.join(bak_dir, fn)
        print(f"  move: {fn}")
        if args.apply:
            shutil.move(src_p, bak_p)

    # 2) ローカル xlsx を NAS にコピー
    local_xlsx = [f for f in os.listdir(src) if f.endswith(".xlsx") and "精算書" in f]
    print(f"\n=== STEP 2: ローカル .xlsx を NAS へコピー ({len(local_xlsx)}件) ===")
    for fn in local_xlsx:
        src_p = os.path.join(src, fn)
        dst_p = os.path.join(dst, fn)
        print(f"  copy: {fn}")
        if args.apply:
            shutil.copy2(src_p, dst_p)

    # 3) 検証
    if args.apply:
        nas_now = sorted([f for f in os.listdir(dst) if f.endswith(".xlsx") and "精算書" in f])
        print(f"\n=== 完了：NAS側 .xlsx は {len(nas_now)} 件 ===")
        for fn in nas_now:
            print(f"  {fn}")
    else:
        print("\n=== dry-run: --apply で実行 ===")


if __name__ == "__main__":
    main()
