@echo off
rem 概要: 「終了時にアプリを自動再起動」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "03_終了時にアプリを自動再起動.bat"
rem 実行結果（想定結果）: 「終了時にアプリを自動再起動」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: Get-Processで対象プロセスを取得、Start-Processでアプリを起動。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。

chcp 65001 >nul
setlocal
set "APP="
set "PROCESS="
set "ARGS="
set /p "APP=監視・起動する exe のフルパス: "
set /p "PROCESS=プロセス名（.exe なし）: "
set /p "ARGS=起動引数（不要なら空欄）: "

if not exist "%APP%" (
    echo 実行ファイルが見つかりません。

    pause
    exit /b 1
)

if "%PROCESS%"=="" exit /b 1
echo 監視中です。終了するには Ctrl+C を押してください。

powershell -NoProfile -Command "while($true){if(-not (Get-Process -Name $env:PROCESS -ErrorAction SilentlyContinue)){Write-Host ('再起動: '+(Get-Date)); Start-Process -FilePath $env:APP -ArgumentList $env:ARGS}; Start-Sleep -Seconds 5}"
