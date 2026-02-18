@echo off
:: このバッチファイルがある場所をアプリのフォルダとして使う
cd /d "%~dp0"

:: Pythonを自動で探す
where python >nul 2>&1
if %errorlevel% == 0 (
    set PYTHON=python
) else (
    where python3 >nul 2>&1
    if %errorlevel% == 0 (
        set PYTHON=python3
    ) else (
        echo Pythonが見つかりませんでした。Pythonをインストールしてください。
        pause
        exit /b 1
    )
)

:: Flaskサーバーをバックグラウンドで起動
start "" /min %PYTHON% app.py

:: サーバーが起動するまで少し待つ
timeout /t 2 /nobreak > nul

:: Chromeをアプリモードで起動
start "" "C:\Program Files\Google\Chrome\Application\chrome.exe" --app=http://localhost:5000 --window-size=1200,800
