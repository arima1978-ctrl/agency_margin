"""送信分xlsmから売上行を抽出する"""
from __future__ import annotations
import os
import warnings
from typing import List, Dict
from datetime import datetime

warnings.filterwarnings("ignore")
import openpyxl

from .config import CATEGORIES, COL_KAZOKU_ID, COL_RYOKIN, COL_JUKUMEI


def _to_int(v):
    if v is None or v == "":
        return None
    try:
        if isinstance(v, str):
            v = v.replace(",", "").strip()
            if v == "":
                return None
        return int(float(v))
    except (TypeError, ValueError):
        return None


def _to_money(v) -> int:
    if v is None or v == "":
        return 0
    try:
        if isinstance(v, str):
            v = v.replace(",", "").strip()
            if v == "":
                return 0
        return int(round(float(v)))
    except (TypeError, ValueError):
        return 0


def extract_sales(xlsm_path: str, target_month: str, nyukin_date: datetime) -> List[Dict]:
    """1つの送信分xlsmから (家族ID, 塾名) ごとに集計したレコード配列を返す。

    各レコードは：
      {家族ID, 塾名, 対象月, 入金日, ④_4カルチャ加盟金…, 速読ID利用料, 合計}
    """
    rows: Dict = {}
    wb = openpyxl.load_workbook(xlsm_path, data_only=True, read_only=True)
    available = set(wb.sheetnames)
    for cat in CATEGORIES:
        if cat not in available:
            continue
        ws = wb[cat]
        for row in ws.iter_rows(min_row=3, values_only=True):
            if row is None:
                continue
            if len(row) <= max(COL_KAZOKU_ID, COL_RYOKIN, COL_JUKUMEI):
                continue
            kid = _to_int(row[COL_KAZOKU_ID])
            if not kid or kid <= 0:
                continue
            ryokin = _to_money(row[COL_RYOKIN])
            if ryokin == 0:
                continue
            juku = row[COL_JUKUMEI]
            juku = juku.strip() if isinstance(juku, str) else (str(juku) if juku else "")
            key = (kid, juku)
            if key not in rows:
                rows[key] = {c: 0 for c in CATEGORIES}
            rows[key][cat] += ryokin
    wb.close()

    out: List[Dict] = []
    for (kid, juku), cats in rows.items():
        rec = {
            "家族ID": kid,
            "塾名": juku,
            "対象月": target_month,
            "入金日": nyukin_date,
            **cats,
            "合計": sum(cats.values()),
        }
        out.append(rec)
    return out


def extract_all(send_specs: List[Dict]) -> List[Dict]:
    """複数送信分を順に抽出してフラット結合。

    send_specs: [{"path", "target_month", "nyukin_date"}, ...]
    """
    all_records = []
    for s in send_specs:
        recs = extract_sales(s["path"], s["target_month"], s["nyukin_date"])
        all_records.extend(recs)
    return all_records
