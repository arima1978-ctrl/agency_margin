"""名簿（カルチャーキッズ名簿.xls）の本部登録シートから家族コード→代理店マップを作る"""
from __future__ import annotations
import warnings
from typing import Tuple, Dict

warnings.filterwarnings("ignore")
import xlrd

from .config import (
    MEIBO_SHEET,
    MEIBO_DATA_START,
    MEIBO_COL_KAZOKU,
    MEIBO_COL_JUKU,
    MEIBO_COL_AGENT,
    AGENT_NAME_NORMALIZE,
)


def _to_int(v):
    if v is None or v == "":
        return None
    try:
        return int(float(v))
    except (TypeError, ValueError):
        return None


def normalize_agent(name: str) -> str:
    if not isinstance(name, str):
        return ""
    name = name.strip()
    return AGENT_NAME_NORMALIZE.get(name, name)


def load_agent_map(meibo_path: str) -> Tuple[Dict[int, str], Dict[int, str]]:
    """名簿を読み、(家族コード→代理店, 家族コード→塾名) を返す。

    同じ家族コードが複数行ある場合、最初に代理店が付いた行を採用。
    代理店が空なら次の行で上書きを試みる。
    """
    book = xlrd.open_workbook(meibo_path)
    sh = book.sheet_by_name(MEIBO_SHEET)
    agent_map: Dict[int, str] = {}
    juku_map: Dict[int, str] = {}
    for r in range(MEIBO_DATA_START, sh.nrows):
        kid = _to_int(sh.cell_value(r, MEIBO_COL_KAZOKU))
        if not kid or kid <= 0:
            continue
        agent = sh.cell_value(r, MEIBO_COL_AGENT)
        agent = normalize_agent(agent) if isinstance(agent, str) else ""
        juku = sh.cell_value(r, MEIBO_COL_JUKU)
        juku = juku.strip() if isinstance(juku, str) else (str(juku) if juku else "")

        if kid not in agent_map or (not agent_map[kid] and agent):
            agent_map[kid] = agent
        if juku and kid not in juku_map:
            juku_map[kid] = juku
    return agent_map, juku_map
