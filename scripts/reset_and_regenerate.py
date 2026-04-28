"""カルチャーキッズマージン明細フォルダを完全リセットし、最新ロジックで再書込みする。

手順:
  1. 既存19ファイル をオリジナル(最古bak)から復元
  2. 新規生成された代理店ファイル＋全bakを削除
  3. 集計実行 → 各代理店ファイルへ書込み

実行例:
  python scripts/reset_and_regenerate.py --apply
  python scripts/reset_and_regenerate.py --dry-run     (削除対象だけ表示)
"""
from __future__ import annotations
import os
import re
import sys
import shutil
import argparse
from collections import defaultdict
from datetime import datetime
from typing import Dict, List

# 親ディレクトリの core/ を読み込めるようにする
HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)

from core.config import EXISTING_AGENT_FILES, DEFAULT_MARGIN_DIR
from core.meibo import load_agent_map
from core.extract import extract_all
from core.aggregate import assign_agent, group_by_agent, UNASSIGNED_LABEL
from core.writer import write_via_excel


FILE_RE = re.compile(r"^カルチャーキッズマージン精算書（(.+)）(_bak_\d{8}_\d{6})?\.xls$")


def scan_dir(margin_dir: str) -> tuple[Dict[str, str], Dict[str, List[str]]]:
    """{agent: main_file_path}, {agent: [bak_paths]} を返す"""
    mains: Dict[str, str] = {}
    baks: Dict[str, List[str]] = defaultdict(list)
    for fn in os.listdir(margin_dir):
        m = FILE_RE.match(fn)
        if not m:
            continue
        agent, bak_marker = m.group(1), m.group(2)
        path = os.path.join(margin_dir, fn)
        if bak_marker:
            baks[agent].append(path)
        else:
            mains[agent] = path
    for k in baks:
        baks[k].sort()  # 名前順 = 時刻順
    return mains, dict(baks)


def restore_originals(mains, baks, dry_run=False) -> List[str]:
    """既存19のうち bak がある場合、最古bakから復元（=オリジナル状態に戻す）"""
    msgs = []
    for agent in EXISTING_AGENT_FILES:
        agent_baks = baks.get(agent, [])
        main_path = mains.get(agent)
        if not agent_baks:
            if main_path:
                msgs.append(f"  [skip] {agent}  bakなし（書込まれていない元のまま）")
            else:
                msgs.append(f"  [warn] {agent}  ファイルが存在しない（要確認）")
            continue
        oldest_bak = agent_baks[0]
        target = main_path or os.path.join(os.path.dirname(oldest_bak),
                                           f"カルチャーキッズマージン精算書（{agent}）.xls")
        msgs.append(f"  [restore] {agent}  ← {os.path.basename(oldest_bak)}  → 上書き")
        if not dry_run:
            shutil.copy2(oldest_bak, target)
            for b in agent_baks:
                os.remove(b)
                msgs.append(f"    rm bak: {os.path.basename(b)}")
    return msgs


def delete_new_agent_files(mains, baks, dry_run=False) -> List[str]:
    """既存19以外の代理店ファイル（main + bak）を全削除"""
    msgs = []
    new_agents = set(mains.keys()) - set(EXISTING_AGENT_FILES)
    new_agents |= (set(baks.keys()) - set(EXISTING_AGENT_FILES))
    for agent in sorted(new_agents):
        if agent in mains:
            msgs.append(f"  [delete] {agent}  → {os.path.basename(mains[agent])}")
            if not dry_run:
                os.remove(mains[agent])
        for b in baks.get(agent, []):
            msgs.append(f"    rm bak: {os.path.basename(b)}")
            if not dry_run:
                os.remove(b)
    return msgs


def run_aggregation(margin_dir: str, send_files: List[Dict], meibo_path: str,
                    sheet_name: str, dry_run=False) -> List[str]:
    msgs = []
    msgs.append(f"  名簿読込: {meibo_path}")
    agent_map, juku_map = load_agent_map(meibo_path)
    msgs.append(f"  名簿の家族コード数: {len(agent_map)}")

    msgs.append("  送信分xlsmから売上抽出（⑮P列で入金確認）…")
    records = extract_all(send_files)
    msgs.append(f"  入金済み売上行: {len(records)}")

    assigned = assign_agent(records, agent_map, juku_map)
    by_agent = group_by_agent(assigned)
    target = {a: r for a, r in by_agent.items() if a != UNASSIGNED_LABEL}
    msgs.append(f"  対象代理店: {len(target)}")

    if dry_run:
        msgs.append("  [dry-run] 書込みは実行しません")
        return msgs

    msgs.append(f"  Excelで書込み開始 (sheet={sheet_name})…")
    results = write_via_excel(
        margin_dir=margin_dir,
        by_agent=target,
        sheet_name=sheet_name,
        backup=False,            # リセット直後なのでbak不要
        create_missing=True,
        progress=lambda i, n, a: print(f"    ... {i}/{n}: {a}"),
    )
    for agent, res in results.items():
        msgs.append(f"    {agent}: {res}")
    return msgs


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--apply", action="store_true", help="実際に変更を加える")
    ap.add_argument("--dry-run", action="store_true", help="変更内容のみ表示")
    ap.add_argument("--margin-dir", default=DEFAULT_MARGIN_DIR)
    ap.add_argument("--sheet-name", default="2026年3月")
    args = ap.parse_args()

    if not args.apply and not args.dry_run:
        ap.error("Specify --apply or --dry-run")
    dry = not args.apply

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

    print(f"==== STEP 1: スキャン ({args.margin_dir}) ====")
    mains, baks = scan_dir(args.margin_dir)
    print(f"  main files: {len(mains)}")
    print(f"  agents with baks: {len(baks)}")

    print(f"\n==== STEP 2: 既存19ファイル復元 (dry={dry}) ====")
    for m in restore_originals(mains, baks, dry_run=dry):
        print(m)

    print(f"\n==== STEP 3: 新規生成ファイル削除 (dry={dry}) ====")
    for m in delete_new_agent_files(mains, baks, dry_run=dry):
        print(m)

    print(f"\n==== STEP 4: 集計再書込み (dry={dry}) ====")
    for m in run_aggregation(args.margin_dir, sends, meibo, args.sheet_name, dry_run=dry):
        print(m)

    print("\n==== DONE ====")


if __name__ == "__main__":
    main()
