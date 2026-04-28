"""代理店ごとの.xlsファイル末尾に新シートを追記する（pywin32経由でExcelを操作）

Excelがインストールされている前提で動く。Excel不在時は AppendToXlsxFallback を使う。
"""
from __future__ import annotations
import os
import sys
import shutil
from typing import Dict, List
from datetime import datetime

from .config import CATEGORIES, LIMITED_AGENT_COLUMNS, EXISTING_AGENT_FILES
from .aggregate import UNASSIGNED_LABEL


def _columns_for_agent(agent: str) -> List[str]:
    return LIMITED_AGENT_COLUMNS.get(agent, CATEGORIES)


def find_agent_file(margin_dir: str, agent: str) -> str | None:
    """代理店名を含むxlsファイルを探す。なければNone"""
    if not os.path.isdir(margin_dir):
        return None
    target = f"カルチャーキッズマージン精算書（{agent}）.xls"
    direct = os.path.join(margin_dir, target)
    if os.path.exists(direct):
        return direct
    # 部分一致で探索
    for fn in os.listdir(margin_dir):
        if not fn.endswith(".xls"):
            continue
        if agent in fn:
            return os.path.join(margin_dir, fn)
    return None


def backup_file(path: str) -> str:
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    bak = path.replace(".xls", f"_bak_{ts}.xls")
    shutil.copy2(path, bak)
    return bak


def write_via_excel(margin_dir: str, by_agent: Dict[str, List[Dict]], sheet_name: str,
                    backup: bool = True, create_missing: bool = True,
                    progress=None) -> Dict[str, str]:
    """Excel COMで各代理店ファイルに新シートを追加。

    Returns: {代理店名: 結果文字列}
    """
    import pythoncom  # type: ignore
    import win32com.client  # type: ignore

    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    results: Dict[str, str] = {}
    try:
        for idx, (agent, records) in enumerate(by_agent.items(), start=1):
            if agent == UNASSIGNED_LABEL:
                results[agent] = "skip:未マッピングのため除外"
                continue
            if progress:
                progress(idx, len(by_agent), agent)

            cols = _columns_for_agent(agent)
            header = ["家族ID", "塾名", "代理店", "対象月", "入金日", *cols, "合計"]

            path = find_agent_file(margin_dir, agent)
            created = False
            if path is None:
                if not create_missing:
                    results[agent] = "skip:ファイルなし"
                    continue
                # 新規ファイル作成（.xls形式）
                path = os.path.join(margin_dir, f"カルチャーキッズマージン精算書（{agent}）.xls")
                wb = excel.Workbooks.Add()
                # デフォルトシート1枚を残す
                ws = wb.Worksheets(1)
                ws.Name = sheet_name
                created = True
            else:
                if backup:
                    backup_file(path)
                wb = excel.Workbooks.Open(os.path.abspath(path))
                # 重複シート名チェック → 既存ありなら _2 を付ける
                existing = {wb.Worksheets(i+1).Name for i in range(wb.Worksheets.Count)}
                target_name = sheet_name
                n = 1
                while target_name in existing:
                    n += 1
                    target_name = f"{sheet_name}_{n}"
                ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
                ws.Name = target_name

            # ヘッダ書込み
            for col_idx, h in enumerate(header, start=1):
                ws.Cells(1, col_idx).Value = h
                ws.Cells(1, col_idx).Font.Bold = True

            # データ書込み
            row_idx = 2
            total_sum = 0
            for r in records:
                row_total = sum(r.get(c, 0) for c in cols)
                ws.Cells(row_idx, 1).Value = r["家族ID"]
                ws.Cells(row_idx, 2).Value = r["塾名"]
                ws.Cells(row_idx, 3).Value = r["代理店"]
                ws.Cells(row_idx, 4).Value = r["対象月"]
                ws.Cells(row_idx, 5).Value = r["入金日"]
                for ci, cat in enumerate(cols, start=6):
                    ws.Cells(row_idx, ci).Value = r.get(cat, 0)
                ws.Cells(row_idx, 5 + len(cols) + 1).Value = row_total
                total_sum += row_total
                row_idx += 1

            row_idx += 1  # 空行
            ws.Cells(row_idx, 5).Value = "売上合計"
            ws.Cells(row_idx, 5).Font.Bold = True
            ws.Cells(row_idx, 5 + len(cols) + 1).Value = total_sum
            ws.Cells(row_idx, 5 + len(cols) + 1).Font.Bold = True

            # 保存（.xls形式維持）
            if created:
                # FileFormat = 56 → xlExcel8 (.xls)
                wb.SaveAs(os.path.abspath(path), FileFormat=56)
                results[agent] = f"created:{os.path.basename(path)}"
            else:
                wb.Save()
                results[agent] = f"updated:{os.path.basename(path)} → sheet '{ws.Name}'"
            wb.Close(SaveChanges=False)

        return results
    finally:
        excel.Quit()
        pythoncom.CoUninitialize()
