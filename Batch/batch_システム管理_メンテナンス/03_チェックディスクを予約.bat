@echo off
rem 概要: 「チェックディスクを予約」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "03_チェックディスクを予約.bat"
rem 実行結果（想定結果）: 「チェックディスクを予約」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: 管理者権限の有無を確認、ディスク検査または修復予約を実行。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
net session >nul 2>&1

if errorlevel 1 (
    echo 管理者として実行してください。

    pause
    exit /b 1
)

set "DRIVE="
set /p "DRIVE=検査するドライブ文字（例: C）: "
if "%DRIVE%"=="" exit /b 1
echo /F と /R でファイルシステム修復と不良セクター検査を指定します。
echo 使用中のドライブでは次回再起動時に実行するか確認されます。
set /p "CONFIRM=続行しますか？ (Y/N): "
if /i not "%CONFIRM%"=="Y" exit /b 0

powershell -NoProfile -Command "$d=$env:DRIVE; if($d -notmatch '^[A-Za-z]$'){Write-Error 'ドライブ文字を1文字で指定してください'; exit 1}; & chkdsk ($d+':') /F /R"

pause
