@echo off
chcp 65001 > nul
cd /d %~dp0

echo ====================================
echo   代理店マージン集計ツールを起動します
echo ====================================
echo.

REM 仮想環境がなければ作成
if not exist .venv (
    echo [初回起動] 仮想環境を作成しています…
    python -m venv .venv
    call .venv\Scripts\activate.bat
    python -m pip install --upgrade pip
    python -m pip install -r requirements.txt
) else (
    call .venv\Scripts\activate.bat
)

REM 念のため最新化
python -m pip install -q -r requirements.txt

echo.
echo ブラウザが開きます。閉じたら、このウィンドウは Ctrl+C で終了してください。
echo.

streamlit run app.py --server.headless false --browser.gatherUsageStats false

pause
