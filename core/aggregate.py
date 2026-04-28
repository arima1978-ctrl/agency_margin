"""売上レコードに代理店を割当て、代理店別にグルーピングする"""
from __future__ import annotations
from collections import defaultdict
from typing import List, Dict, Tuple

from .config import CATEGORIES
from .meibo import normalize_agent

UNASSIGNED_LABEL = "(未設定)"


def assign_agent(records: List[Dict], agent_map: Dict[int, str], juku_map: Dict[int, str]) -> List[Dict]:
    """各レコードに代理店を付与する（コピー、元配列は変更しない）。

    塾名が空の場合は名簿から補完する。
    """
    out = []
    for r in records:
        kid = r["家族ID"]
        agent = agent_map.get(kid, "") or ""
        agent = normalize_agent(agent)
        if not agent:
            agent = UNASSIGNED_LABEL
        juku = r.get("塾名") or ""
        if not juku and kid in juku_map:
            juku = juku_map[kid]
        new = {**r, "代理店": agent, "塾名": juku}
        out.append(new)
    return out


def group_by_agent(records: List[Dict]) -> Dict[str, List[Dict]]:
    """代理店ごとにグルーピング（売上合計の降順で代理店をソート）"""
    by_agent: Dict[str, List[Dict]] = defaultdict(list)
    for r in records:
        by_agent[r["代理店"]].append(r)
    # 各グループを (対象月, 家族ID) 順にソート
    for k in by_agent:
        by_agent[k].sort(key=lambda r: (r["対象月"], r["家族ID"]))
    return dict(by_agent)


def agent_totals(by_agent: Dict[str, List[Dict]]) -> List[Tuple[str, int, int]]:
    """[(代理店, 件数, 売上合計), ...] を売上降順で返す"""
    rows = []
    for agent, recs in by_agent.items():
        rows.append((agent, len(recs), sum(r["合計"] for r in recs)))
    rows.sort(key=lambda x: x[2], reverse=True)
    return rows
