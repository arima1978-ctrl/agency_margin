"""カルチャーキッズ代理店マージン集計 — Streamlit UI

事務員が3か月に1回、3つの送信分xlsm + 名簿 を投入して
代理店ファイルに売上シートを追記するためのインターフェース。
"""
from __future__ import annotations
import os
import io
import sys
import tempfile
from datetime import datetime, date

import streamlit as st

# ローカルパッケージ
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from core.config import (
    DEFAULT_MARGIN_DIR,
    PREVIEW_FILENAME_TEMPLATE,
    CATEGORIES,
)
from core.meibo import load_agent_map
from core.extract import extract_all
from core.aggregate import assign_agent, group_by_agent, agent_totals, UNASSIGNED_LABEL
from core.preview import write_preview


st.set_page_config(page_title="代理店マージン集計", page_icon="📊", layout="wide")
st.title("📊 カルチャーキッズ 代理店マージン集計")
st.caption("3か月に1回、3つの送信分xlsm＋名簿から代理店ごとに売上を集計します。")

# ---------- 1. ファイルアップロード ----------
st.header("1. ファイルを選択")

col_a, col_b = st.columns([1, 2])
with col_a:
    st.markdown("**送信分xlsm（3つ）**")
with col_b:
    st.markdown("各送信分の `(入金チェック）.xlsm` をアップロードしてください。")

xlsm_files = st.file_uploader(
    "送信分xlsm（3ファイル）",
    type=["xlsm", "xlsx"],
    accept_multiple_files=True,
    key="xlsm",
)

meibo_file = st.file_uploader(
    "名簿（カルチャーキッズ名簿.xls）",
    type=["xls", "xlsx"],
    key="meibo",
)


def _save_uploaded(uf) -> str:
    """アップロードされたファイルを一時ファイルに保存しパスを返す"""
    suffix = os.path.splitext(uf.name)[1]
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    tmp.write(uf.getbuffer())
    tmp.close()
    return tmp.name


# ---------- 2. 対象月 / 入金日設定 ----------
st.header("2. 各送信分の対象月・入金日を設定")
st.caption("送信分のファイル名から自動推定しています。違っていたら修正してください。")


def _infer_target_month(filename: str) -> tuple[str, date]:
    """ファイル名から対象月と入金日を推定する"""
    import re
    m = re.search(r"(\d{4})年(\d{1,2})月", filename)
    if m:
        y, mo = int(m.group(1)), int(m.group(2))
        # 送信月の前月が対象月
        target_y, target_m = (y, mo - 1) if mo > 1 else (y - 1, 12)
        # 入金日は送信月の翌月6日とする
        nyu_y, nyu_m = (y, mo)
        return f"{target_y}年{target_m:02d}月", date(nyu_y, nyu_m, 6)
    return "", date.today()


send_specs = []
if xlsm_files:
    sorted_xlsm = sorted(xlsm_files, key=lambda f: f.name)
    for i, f in enumerate(sorted_xlsm[:3]):
        target_default, nyukin_default = _infer_target_month(f.name)
        c1, c2, c3 = st.columns([3, 2, 2])
        with c1:
            st.text(f"📄 {f.name}")
        with c2:
            target = st.text_input(
                "対象月", value=target_default, key=f"target_{i}"
            )
        with c3:
            nyukin = st.date_input(
                "入金日", value=nyukin_default, key=f"nyukin_{i}"
            )
        send_specs.append({"file": f, "target_month": target, "nyukin_date": nyukin})

# ---------- 3. 出力設定 ----------
st.header("3. 出力先・追加シート名")
c1, c2 = st.columns([2, 1])
with c1:
    margin_dir = st.text_input(
        "代理店マージン明細フォルダ",
        value=DEFAULT_MARGIN_DIR,
        help="19代理店ファイルの置場。新規代理店ファイルもここに作成します。",
    )
with c2:
    sheet_name = st.text_input(
        "追加シート名",
        value=f"{datetime.now().year}年{((datetime.now().month - 1)//3 + 1)*3}月",
        help="各代理店xlsに追加する新シートの名前",
    )


# ---------- 4. プレビュー実行 ----------
st.header("4. 集計プレビュー")
run_preview = st.button("📊 集計プレビューを生成", type="primary", disabled=not(xlsm_files and meibo_file and len(send_specs) >= 1))

if run_preview:
    with st.spinner("名簿を読み込んでいます…"):
        meibo_path = _save_uploaded(meibo_file)
        agent_map, juku_map = load_agent_map(meibo_path)
    st.write(f"名簿の家族コード数: **{len(agent_map)}**")

    with st.spinner("送信分xlsmから売上を抽出中…"):
        for s in send_specs:
            s["path"] = _save_uploaded(s["file"])
            s["nyukin_date"] = datetime.combine(s["nyukin_date"], datetime.min.time())
        all_records = extract_all(send_specs)
    st.write(f"抽出した売上行: **{len(all_records)}**")

    with st.spinner("代理店割当・集計中…"):
        assigned = assign_agent(all_records, agent_map, juku_map)
        by_agent = group_by_agent(assigned)
    st.session_state["by_agent"] = by_agent
    st.session_state["sheet_name"] = sheet_name
    st.session_state["margin_dir"] = margin_dir

# 結果表示（セッションステート利用）
if "by_agent" in st.session_state:
    by_agent = st.session_state["by_agent"]
    # 代理店ありのみカウント（未設定=直販想定で対象外）
    mapped_agents = [a for a in by_agent if a != UNASSIGNED_LABEL]
    unassigned_count = len(by_agent.get(UNASSIGNED_LABEL, []))
    st.success(f"代理店数: **{len(mapped_agents)}**（直販＝未マッピング {unassigned_count} 行は対象外）")

    # 対応表 — 未設定は除外
    rows = [r for r in agent_totals(by_agent) if r[0] != UNASSIGNED_LABEL]
    st.subheader("📋 対応表（代理店別）")
    import pandas as pd  # streamlit同梱
    df_tot = pd.DataFrame(rows, columns=["代理店", "件数", "売上合計"])
    st.dataframe(df_tot, use_container_width=True, hide_index=True)

    # 各代理店プレビュー（タブ表示）— 未設定は除外
    st.subheader("🏢 代理店別プレビュー")
    agent_names = [r[0] for r in rows]
    if agent_names:
        tab_objs = st.tabs([f"{a} ({len(by_agent[a])})" for a in agent_names])
        for tab, agent in zip(tab_objs, agent_names):
            with tab:
                df = pd.DataFrame(by_agent[agent])
                st.dataframe(df, use_container_width=True, hide_index=True)

    # 直販（未マッピング）— 確認用に折りたたみで表示
    if unassigned_count:
        with st.expander(f"⚠️ 直販／代理店未設定 {unassigned_count} 行（参考表示・対象外）"):
            df_un = pd.DataFrame(by_agent[UNASSIGNED_LABEL])
            st.dataframe(df_un, use_container_width=True, hide_index=True)

    # プレビューxlsxダウンロード
    st.subheader("⬇️ プレビューファイル")
    out_dir = tempfile.mkdtemp(prefix="agency_margin_")
    out_path = os.path.join(out_dir, PREVIEW_FILENAME_TEMPLATE.format(quarter=st.session_state["sheet_name"]))
    write_preview(out_path, by_agent, st.session_state["sheet_name"])
    with open(out_path, "rb") as f:
        st.download_button(
            "📥 プレビューxlsxをダウンロード",
            data=f.read(),
            file_name=os.path.basename(out_path),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    # ---------- 5. 実書込み ----------
    st.header("5. 各代理店ファイルに書き込み")
    st.warning(
        "この操作は各代理店の `.xls` ファイルに新シートを追加します。"
        "実行前に **必ずプレビューを確認** してください。"
        " バックアップは自動で同フォルダに作成されます（`*_bak_YYYYMMDD_HHMMSS.xls`）。"
    )
    cwc1, cwc2 = st.columns(2)
    with cwc1:
        do_backup = st.checkbox("バックアップ作成", value=True)
        skip_unassigned = st.checkbox("未マッピングはスキップ", value=True)
    with cwc2:
        create_missing = st.checkbox("代理店ファイルが無ければ新規作成", value=True)

    if st.button("✅ 各代理店ファイルに書込み実行", type="primary"):
        try:
            from core.writer import write_via_excel

            target_by_agent = dict(by_agent)
            if skip_unassigned:
                target_by_agent.pop(UNASSIGNED_LABEL, None)

            with st.spinner("Excelを起動して書込み中…（数十秒かかる場合があります）"):
                progress_bar = st.progress(0)
                msg = st.empty()

                def _progress(i, n, agent):
                    progress_bar.progress(i / max(n, 1))
                    msg.text(f"処理中 {i}/{n}: {agent}")

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
            st.success("書込み完了")
            df_res = pd.DataFrame(
                [(k, v) for k, v in results.items()],
                columns=["代理店", "結果"],
            )
            st.dataframe(df_res, use_container_width=True, hide_index=True)
        except ImportError as e:
            st.error(f"pywin32が必要です: {e}")
        except Exception as e:
            st.error(f"書込み中にエラーが発生しました: {e}")
            st.exception(e)

# ---------- フッター ----------
st.divider()
st.caption("agency_margin v0.1 — github.com/arima1978-ctrl/agency_margin")
