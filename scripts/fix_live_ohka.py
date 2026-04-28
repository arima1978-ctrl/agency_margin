"""s-Live.xls / 桜花名塾植田校.xls を削除し、その分のデータを Live.xls / 桜花名塾.xls に追記する。

前回 reset_and_regenerate.py で Live と s-Live、桜花名塾と桜花名塾植田校 を別ファイルとして
書き込んでしまったので、正規化を反転（s-Live→Live, 桜花名塾植田校→桜花名塾）したうえで
履歴のあるファイル側に統合する。
"""
from __future__ import annotations
import os, sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)

from core.config import DEFAULT_MARGIN_DIR
from core.meibo import load_agent_map
from core.extract import extract_all
from core.aggregate import assign_agent, group_by_agent, UNASSIGNED_LABEL
from core.writer import write_via_excel


def main():
    margin_dir = DEFAULT_MARGIN_DIR
    sheet_name = "2026年3月"

    # 不要ファイルを削除
    for fn in [
        "カルチャーキッズマージン精算書（s-Live）.xls",
        "カルチャーキッズマージン精算書（桜花名塾植田校）.xls",
    ]:
        p = os.path.join(margin_dir, fn)
        if os.path.exists(p):
            os.remove(p)
            print(f"  rm: {fn}")
        else:
            print(f"  (skip, not found): {fn}")

    # 集計再実行（正規化反転後）
    BASE = r"C:\Users\USER\Documents\三浦さんマージン清算"
    sends = [
        {"path": os.path.join(BASE, "2025年12月17日送信分", "2025年12月18日送信(入金チェック）.xlsm"),
         "target_month": "2025年11月"},
        {"path": os.path.join(BASE, "2026年1月24日送信分", "2026年1月24送信(入金チェック）.xlsm"),
         "target_month": "2025年12月"},
        {"path": os.path.join(BASE, "2026年2月17日送信分", "2026年2月17日送信(入金チェック）.xlsm"),
         "target_month": "2026年1月"},
    ]
    meibo = os.path.join(BASE, "カルチャーキッズ名簿.xls")

    agent_map, juku_map = load_agent_map(meibo)
    records = extract_all(sends)
    assigned = assign_agent(records, agent_map, juku_map)
    by_agent = group_by_agent(assigned)

    # Live と 桜花名塾 だけ抽出
    target = {a: r for a, r in by_agent.items() if a in {"Live", "桜花名塾"}}
    print(f"  対象代理店: {list(target.keys())}")
    for a, recs in target.items():
        print(f"    {a}: {len(recs)}件 売上={sum(r['合計'] for r in recs):,}")

    if not target:
        print("  対象データなし")
        return

    # 既存履歴ファイルへ追記（バックアップ作成）
    print(f"  Excelで書込み (sheet={sheet_name})…")
    results = write_via_excel(
        margin_dir=margin_dir,
        by_agent=target,
        sheet_name=sheet_name,
        backup=True,
        create_missing=False,
        progress=lambda i, n, a: print(f"    ... {i}/{n}: {a}"),
    )
    for a, res in results.items():
        print(f"    {a}: {res}")
    print("DONE")


if __name__ == "__main__":
    main()
