@echo off
rem 概要: 「Slack互換Webhookで通知」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "19_Slack互換Webhookで通知.bat"
rem 実行結果（想定結果）: 「Slack互換Webhookで通知」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: ConvertTo-JsonでJSON形式に変換、Invoke-RestMethodでHTTP APIへ要求を送信。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal

if not defined CHAT_WEBHOOK_URL (
    echo CHAT_WEBHOOK_URL 環境変数にWebhook URLを設定してください。
    echo URLをソースコードへ直接書かないでください。
    exit /b 1
)

set "MESSAGE="
set /p "MESSAGE=送信するメッセージ: "
if "%MESSAGE%"=="" exit /b 1

powershell -NoProfile -Command "if($env:CHAT_WEBHOOK_URL -notmatch '^https://'){throw 'HTTPS の webhook URL が必要です'}; $body=@{text=$env:MESSAGE} | ConvertTo-Json -Compress; Invoke-RestMethod -Uri $env:CHAT_WEBHOOK_URL -Method Post -ContentType 'application/json' -Body $body -ErrorAction Stop | Out-Null"

if errorlevel 1 (
    echo Webhook 通知に失敗しました。URLと接続先のペイロード形式を確認してください。
    exit /b 1
)

echo 通知を送信しました。
