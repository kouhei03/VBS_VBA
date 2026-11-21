@echo off
rem 概要: 「指定アプリを終了してゲーム起動」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "07_指定アプリを終了してゲーム起動.bat"
rem 実行結果（想定結果）: 「指定アプリを終了してゲーム起動」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: Get-Processで対象プロセスを取得、Stop-Processで指定プロセスを停止。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set "PROCESSES="
set "GAME="
set /p "PROCESSES=終了するアプリ名をカンマ区切りで入力（.exe なし）: "
set /p "GAME=起動するゲームの exe フルパス: "

if not exist "%GAME%" (
    echo ゲームの実行ファイルが見つかりません。

    pause
    exit /b 1
)

echo 入力したアプリのみを強制終了します。未保存データが失われる場合があります。
set "CONFIRM="
set /p "CONFIRM=続行する場合は CLOSE と入力してください: "
if not "%CONFIRM%"=="CLOSE" exit /b 0

powershell -NoProfile -Command "$names=@($env:PROCESSES -split ',' | ForEach-Object {$_.Trim()} | Where-Object {$_}); foreach($n in $names){Get-Process -Name $n -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction Stop}"

if errorlevel 1 (
    echo 指定アプリを終了できませんでした。ゲームは起動しません。

    pause
    exit /b 1
)

start "" "%GAME%"
