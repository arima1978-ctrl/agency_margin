"""カルチャーキッズ代理店マージン集計 — Streamlit UI

事務員が3か月に1回、3つの送信分xlsm + 名簿 から代理店ファイルに売上シートを
追記するためのインターフェース。フォルダを指定するだけで全自動検出する。
"""
from __future__ import annotations
import os
import re
import sys
import shutil
import tempfile
import subprocess
from datetime import datetime
from typing import List, Dict, Optional

import streamlit as st
import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from core.config import (
    DEFAULT_MARGIN_DIR,
    PREVIEW_FILENAME_TEMPLATE,
)
from core.meibo import load_agent_map
from core.extract import extract_all
from core.aggregate import assign_agent, group_by_agent, agent_totals, UNASSIGNED_LABEL
from core.preview import write_preview


# ---------- ページ設定 ----------
st.set_page_config(page_title="代理店マージン集計", page_icon="📊", layout="wide")

# ---------- サイドバー（操作ガイド） ----------
with st.sidebar:
    st.markdown("### 📖 使い方")
    st.markdown(
        """
1. **親フォルダ**（基準パス）を入力
2. **送信分xlsm 3つ＋名簿** のフルパスを入力
3. **代理店マージン明細フォルダ** を確認
4. **集計実行** → プレビューで内容確認
5. **書込み** → 各代理店ファイルに追記
        """
    )
    st.divider()
    st.markdown("### 💡 ヒント")
    st.markdown(
        """
- 入金日は ⑮入金チェックシートのP列から自動取得
- P列が空欄の家族IDは未入金として除外
- 既存ファイルはバックアップを自動作成
- 代理店ファイルが無い場合は自動で新規作成
        """
    )
    st.divider()
    st.markdown("### 🛠️ トラブル時")
    st.markdown(
        """
- 検出されない → ファイル名が `YYYY年MM月DD日送信分` で始まるか確認
- Excel書込み失敗 → Excelが起動中でないか確認
- 未マッピング多い → 名簿の代理店列を確認
        """
    )
    st.divider()
    if st.button("🔄 セッションリセット"):
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.rerun()

# ---------- メイン ----------
st.title("📊 カルチャーキッズ 代理店マージン集計")
st.caption("3か月に1回、3つの送信分xlsm＋名簿から代理店ごとに売上を集計します。")


# ============== STEP 1: フォルダ選択 ==============
st.header("STEP 1：データフォルダを選択")

from core.config import DEFAULT_PARENT_DIR

parent_dir = st.text_input(
    "三浦さんマージン清算フォルダ",
    value=DEFAULT_PARENT_DIR,
    help="送信分フォルダ・名簿・代理店マージン明細を含む親フォルダのパス",
)


def infer_target_month_from_path(path: str) -> str:
    """ファイル名・パスから "YYYY年MM月" 推定。送信YYYY年MM月の前月を対象月とする"""
    m = re.search(r"(\d{4})年(\d{1,2})月", os.path.basename(path))
    if not m:
        m = re.search(r"(\d{4})年(\d{1,2})月", path)
    if not m:
        return ""
    y, mo = int(m.group(1)), int(m.group(2))
    target_y, target_m = (y, mo - 1) if mo > 1 else (y - 1, 12)
    return f"{target_y}年{target_m:02d}月"


# ============== STEP 2: 集計対象を手動で指定 ==============
st.header("STEP 2：集計対象を手動で指定")

st.markdown("**送信分xlsm を 3 つ＋名簿を、フルパスで入力してください**（NAS上のパスをそのままコピペ可）")
st.caption(
    "例：`" + os.path.join(parent_dir, "2026年2月17日送信分", "2026年2月17日送信(入金チェック）.xlsm") + "`"
)

ok = True
sends_paths = []
sends_targets = []

for i in range(3):
    cols = st.columns([3, 1])
    with cols[0]:
        path = st.text_input(
            f"送信分 {i+1} のフルパス（.xlsm）",
            key=f"send_path_{i}",
            placeholder="/mnt/nas_share/.../○○送信分/○○送信(入金チェック）.xlsm  または  C:\\Users\\...",
        )
    with cols[1]:
        # 自動推定された対象月をデフォルトに、手動修正可
        default_tm = infer_target_month_from_path(path) if path else ""
        target_month = st.text_input(
            f"対象月 {i+1}",
            value=default_tm,
            key=f"send_tm_{i}",
            placeholder="2026年01月",
        )
    if path:
        if not os.path.isfile(path):
            st.error(f"  ❌ ファイルが見つかりません: {path}")
            ok = False
        elif not path.lower().endswith((".xlsm", ".xlsx")):
            st.error(f"  ❌ xlsm/xlsx ファイルを指定してください: {os.path.basename(path)}")
            ok = False
        elif not target_month.strip():
            st.warning(f"  ⚠️ 対象月を入力してください")
            ok = False
        else:
            st.success(f"  ✅ {os.path.basename(path)}  →  対象月 {target_month}")
            sends_paths.append(path)
            sends_targets.append(target_month.strip())
    else:
        ok = False

# 名簿
st.divider()
meibo_path = st.text_input(
    "名簿ファイルのフルパス（.xls / .xlsx）",
    placeholder="/mnt/nas_share/.../カルチャーキッズ名簿.xls",
)
if meibo_path:
    if not os.path.isfile(meibo_path):
        st.error(f"❌ 名簿が見つかりません: {meibo_path}")
        ok = False
    else:
        st.success(f"✅ 名簿: {os.path.basename(meibo_path)}")
else:
    ok = False

# 代理店マージン明細フォルダ
st.divider()
margin_dir = st.text_input(
    "代理店マージン明細フォルダ（出力先）",
    value=os.path.join(parent_dir, "カルチャーキッズマージン明細"),
)
if margin_dir and os.path.isdir(margin_dir):
    n_files = sum(1 for f in os.listdir(margin_dir)
                  if f.endswith((".xlsx", ".xls")) and "_bak_" not in f and "精算書" in f)
    st.success(f"✅ 出力先: `{margin_dir}`（既存{n_files}ファイル）")
else:
    st.error(f"❌ フォルダが見つかりません: {margin_dir}")
    ok = False

# ============== STEP 3: 集計実行 ==============
st.header("STEP 3：集計設定")

c1, c2 = st.columns(2)
with c1:
    # 追加シート名のデフォルト = 入力された対象月のうち最新の四半期末
    default_sheet = ""
    if sends_targets:
        # "2026年01月" などをパースして最新を取る
        latest_y, latest_m = 0, 0
        for tm in sends_targets:
            m = re.match(r"(\d{4})年(\d{1,2})月", tm)
            if m:
                y, mo = int(m.group(1)), int(m.group(2))
                if (y, mo) > (latest_y, latest_m):
                    latest_y, latest_m = y, mo
        if latest_y:
            target_q_y, target_q_m = (latest_y, latest_m + 2) if latest_m + 2 <= 12 else (latest_y + 1, latest_m + 2 - 12)
            # シンプルに 入金月＝対象月+1 の四半期末で
            default_sheet = f"{target_q_y}年{((target_q_m - 1)//3 + 1)*3}月"
    if not default_sheet:
        default_sheet = f"{datetime.now().year}年{((datetime.now().month - 1)//3 + 1)*3}月"
    sheet_name = st.text_input(
        "追加シート名",
        value=default_sheet,
        help="各代理店xlsxに追加する新シートの名前（通常は四半期末月）",
    )
with c2:
    st.markdown("**自動取得**")
    st.markdown("- 入金日: 各xlsmの ⑮入金チェックシートP列")
    st.markdown("- 対象月: 送信日の前月")

# 集計実行ボタン
run_btn = st.button(
    "📊 集計を実行",
    type="primary",
    disabled=not ok,
    use_container_width=True,
)


def _save_uploaded(uf) -> str:
    suffix = os.path.splitext(uf.name)[1]
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    tmp.write(uf.getbuffer())
    tmp.close()
    return tmp.name


if run_btn:
    with st.spinner("名簿を読み込んでいます…"):
        agent_map, juku_map = load_agent_map(meibo_path)
    st.write(f"名簿の家族コード数: **{len(agent_map):,}**")

    with st.spinner("送信分xlsmから売上を抽出中…（⑮入金チェックシート参照）"):
        send_specs = [
            {"path": p, "target_month": tm}
            for p, tm in zip(sends_paths, sends_targets)
        ]
        all_records = extract_all(send_specs)
    st.write(f"入金済み売上行: **{len(all_records):,}**")

    with st.spinner("代理店割当・集計中…"):
        assigned = assign_agent(all_records, agent_map, juku_map)
        by_agent = group_by_agent(assigned)
    st.session_state["by_agent"] = by_agent
    st.session_state["sheet_name"] = sheet_name
    st.session_state["margin_dir"] = margin_dir

# ============== STEP 4: 結果プレビュー ==============
if "by_agent" in st.session_state:
    st.divider()
    st.header("STEP 4：結果プレビュー")

    by_agent = st.session_state["by_agent"]
    mapped_agents = [a for a in by_agent if a != UNASSIGNED_LABEL]
    unassigned_count = len(by_agent.get(UNASSIGNED_LABEL, []))
    rows = [r for r in agent_totals(by_agent) if r[0] != UNASSIGNED_LABEL]

    cm1, cm2, cm3 = st.columns(3)
    with cm1:
        st.metric("代理店数", f"{len(mapped_agents)}")
    with cm2:
        st.metric("売上行合計", f"{sum(r[1] for r in rows):,}")
    with cm3:
        total_yen = sum(r[2] for r in rows)
        st.metric("売上合計金額", f"¥{total_yen:,}")

    st.subheader("📋 代理店別 集計表")
    df_tot = pd.DataFrame(rows, columns=["代理店", "件数", "売上合計"])
    st.dataframe(df_tot, use_container_width=True, hide_index=True)

    with st.expander("🏢 代理店別の明細を見る"):
        agent_names = [r[0] for r in rows]
        if agent_names:
            tab_objs = st.tabs([f"{a} ({len(by_agent[a])})" for a in agent_names])
            for tab, agent in zip(tab_objs, agent_names):
                with tab:
                    df = pd.DataFrame(by_agent[agent])
                    st.dataframe(df, use_container_width=True, hide_index=True)

    if unassigned_count:
        with st.expander(f"⚠️ 直販／代理店未設定 {unassigned_count} 行（参考表示・対象外）"):
            df_un = pd.DataFrame(by_agent[UNASSIGNED_LABEL])
            st.dataframe(df_un, use_container_width=True, hide_index=True)

    # プレビューxlsxダウンロード
    st.subheader("📥 プレビューファイル（事前チェック用）")
    out_dir = tempfile.mkdtemp(prefix="agency_margin_")
    out_path = os.path.join(out_dir, PREVIEW_FILENAME_TEMPLATE.format(quarter=st.session_state["sheet_name"]))
    write_preview(out_path, by_agent, st.session_state["sheet_name"])
    with open(out_path, "rb") as f:
        st.download_button(
            "プレビューxlsxをダウンロード",
            data=f.read(),
            file_name=os.path.basename(out_path),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    # ============== STEP 5: 書込み ==============
    st.divider()
    st.header("STEP 5：各代理店ファイルへ書込み")
    st.info(
        "✅ 各代理店の `.xls` に「" + st.session_state["sheet_name"] + "」シートを追加します。\n\n"
        "📦 既存ファイルは自動でバックアップされます（`*_bak_YYYYMMDD_HHMMSS.xls`）。\n\n"
        "🆕 代理店ファイルが無い場合は同フォルダに新規作成されます。"
    )

    cwc1, cwc2 = st.columns(2)
    with cwc1:
        do_backup = st.checkbox("バックアップ作成", value=True)
    with cwc2:
        create_missing = st.checkbox("無い代理店ファイルは新規作成", value=True)

    if st.button("✅ 書込み実行", type="primary", use_container_width=True):
        try:
            from core.writer import write_via_excel

            target_by_agent = {a: r for a, r in by_agent.items() if a != UNASSIGNED_LABEL}

            progress_bar = st.progress(0)
            msg = st.empty()

            def _progress(i, n, agent):
                progress_bar.progress(i / max(n, 1))
                msg.text(f"処理中 {i}/{n}: {agent}")

            with st.spinner("Excelを起動して書込み中…"):
                results = write_via_excel(
                    margin_dir=st.session_state["margin_dir"],
                    by_agent=target_by_agent,
                    sheet_name=st.session_state["sheet_name"],
                    backup=do_backup,
                    create_missing=create_missing,
                    progress=_progress,
                )
                progress_bar.progress(1.0)
                msg.empty()

            st.success("✅ 書込み完了！")

            df_res = pd.DataFrame(
                [(k, v) for k, v in results.items()],
                columns=["代理店", "結果"],
            )
            st.dataframe(df_res, use_container_width=True, hide_index=True)

            # 完了後ボタン
            cb1, cb2 = st.columns(2)
            with cb1:
                if st.button("📁 マージン明細フォルダを開く"):
                    try:
                        os.startfile(st.session_state["margin_dir"])
                    except Exception as e:
                        st.error(f"フォルダを開けませんでした: {e}")
            with cb2:
                st.markdown(f"📍 出力先: `{st.session_state['margin_dir']}`")

        except ImportError as e:
            st.error(f"pywin32が必要です: {e}")
        except Exception as e:
            st.error(f"書込み中にエラーが発生しました: {e}")
            st.exception(e)


# ---------- フッター ----------
st.divider()
st.caption("agency_margin v0.2 — github.com/arima1978-ctrl/agency_margin")
