# agency_margin

カルチャーキッズ代理店マージン集計ツール。3か月に1回、3つの送信分xlsm＋名簿から代理店ごとに売上を集計し、各代理店ファイルに新シートとして追記する。

## 起動方法

### 事務員向け（簡単起動）

1. このフォルダの `start.bat` をダブルクリック
2. ブラウザが自動で開く（http://localhost:8501）
3. 画面の手順に従う

### 開発者

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

## 操作手順

1. **ファイルを選択**
   - 送信分xlsm（3つ）：例 `2025年12月18日送信(入金チェック）.xlsm`
   - 名簿：`カルチャーキッズ名簿.xls`
2. **対象月・入金日を確認**（ファイル名から自動推定。違っていたら修正）
3. **出力先・シート名を設定**
4. **集計プレビューを生成** → 代理店別タブで内容確認
5. **プレビューxlsxをダウンロード**して目視確認
6. 問題なければ **各代理店ファイルに書込み実行**

## ディレクトリ構成

```
agency_margin/
├── app.py                 # Streamlit UI
├── core/
│   ├── config.py          # カテゴリ・正規化辞書・列インデックス
│   ├── meibo.py           # 名簿パース
│   ├── extract.py         # xlsm売上抽出
│   ├── aggregate.py       # 代理店ごと集計
│   ├── preview.py         # プレビューxlsx生成
│   └── writer.py          # 各代理店xlsへ追記（pywin32経由）
├── start.bat              # 事務員向けダブルクリック起動
├── requirements.txt
└── README.md
```

## 集計ロジック

- 各送信分xlsmの **8つのカテゴリシート**（`④_4カルチャ加盟金` ほか）から `家族ID / 料金 / 塾名` を抽出
- 名簿の `本部登録` シートで `家族コード → 代理店` を引いて割当
- 同じ家族IDの売上を1行に横展開（カテゴリ列ごと）
- 19代理店ファイルのうち、**朝日教育社・中央教育研究所・誠伸社** は限定列フォーマットを保持

## 表記揺れ対応

`core/config.py` の `AGENT_NAME_NORMALIZE` 辞書に登録：

```python
AGENT_NAME_NORMALIZE = {
    "朝日教育": "朝日教育社",
    # 必要に応じて追加
}
```

## 安全策

- 既存ファイルは自動でバックアップ（`*_bak_YYYYMMDD_HHMMSS.xls`）
- シート追加のみ（既存シートの書換えなし）
- 未マッピング行は別タブで表示し、書込み時はスキップ可能

## 動作要件

- Windows + Excel（`.xls`書込みのため `pywin32` 経由でExcel COMを利用）
- Python 3.10+
