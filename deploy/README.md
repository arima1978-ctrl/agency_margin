# Linux サーバー（192.168.1.16）デプロイ手順

社内NAS（192.168.1.201）上のファイルを直接読み書きする Streamlit を、
192.168.1.16:8506 で常時稼働させる。

## 前提
- skyuser アカウント
- NAS が `/mnt/nas_share` にマウント済み（cifs, guest）
- Python 3.13+

## 1. クローン＆仮想環境

```bash
cd ~
git clone https://github.com/arima1978-ctrl/agency_margin.git
cd agency_margin
python3 -m venv .venv
.venv/bin/pip install --upgrade pip
.venv/bin/pip install -r requirements.txt
```

## 2. systemd ユニット

```bash
sudo cp deploy/agency_margin.service /etc/systemd/system/agency_margin.service
sudo systemctl daemon-reload
sudo systemctl enable agency_margin
sudo systemctl start agency_margin
sudo systemctl status agency_margin
```

## 3. 動作確認

社内LANから http://192.168.1.16:8506 にアクセス。

## 4. ログ確認

```bash
sudo journalctl -u agency_margin -f
```

## 5. 更新

```bash
cd ~/agency_margin
git pull
.venv/bin/pip install -r requirements.txt
sudo systemctl restart agency_margin
```

## 環境変数

- `AGENCY_MARGIN_PARENT`: 三浦さんマージン清算 親フォルダ
- `AGENCY_MARGIN_DIR`: カルチャーキッズマージン明細 フォルダ
