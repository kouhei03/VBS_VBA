@echo off
rem 概要: 「常駐アプリを終了して再起動」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "10_常駐アプリを終了して再起動.bat"
rem 実行結果（想定結果）: 「常駐アプリを終了して再起動」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: 指定プロセスを終了、指定秒数の待機を実行。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set "PROCESS="
set "EXE_PATH="
set /p "PROCESS=終了するプロセス名（.exe なし）: "
set /p "EXE_PATH=再起動するアプリの exe フルパス: "
if "%PROCESS%"=="" exit /b 1

if not exist "%EXE_PATH%" (
    echo 実行ファイルが見つかりません。

    pause
    exit /b 1
)

echo 強制終了すると未保存データが失われる可能性があります。
set "CONFIRM="
set /p "CONFIRM=続行する場合は RESTART と入力してください: "
if not "%CONFIRM%"=="RESTART" exit /b 0

taskkill /F /IM "%PROCESS%.exe"

timeout /t 2 /nobreak >nul
start "" "%EXE_PATH%"
if errorlevel 1 echo アプリを起動できませんでした。

pause
