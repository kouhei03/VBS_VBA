@echo off
rem 概要: 「バックアップ後にWebhook通知」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "20_バックアップ後にWebhook通知.bat"
rem 実行結果（想定結果）: 「バックアップ後にWebhook通知」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: Robocopyでファイルをコピーまたは同期、ConvertTo-JsonでJSON形式に変換、Invoke-RestMethodでHTTP APIへ要求を送信。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set /p "SOURCE=バックアップ元フォルダ: "
set /p "DEST=バックアップ先フォルダ: "
set /p "WEBHOOK=通知先 Webhook URL（秘密情報として扱ってください）: "

if not exist "%SOURCE%\." (
    echo バックアップ元が見つかりません。

    pause
    exit /b 1
)

robocopy "%SOURCE%" "%DEST%" /E /Z /R:3 /W:5 /XJ
set "RC=%ERRORLEVEL%"
set "RESULT=成功"
if %RC% GEQ 8 set "RESULT=失敗"

powershell -NoProfile -Command "$payload=@{text=('バックアップ'+$env:RESULT+'。Robocopy 終了コード: '+$env:RC)} | ConvertTo-Json; Invoke-RestMethod -Uri $env:WEBHOOK -Method Post -ContentType 'application/json; charset=utf-8' -Body ([Text.Encoding]::UTF8.GetBytes($payload))"
if errorlevel 1 echo 通知に失敗しました。Webhook の仕様や URL を確認してください。

if %RC% GEQ 8 (
    echo バックアップ失敗。終了コード: %RC%
    exit /b %RC%
)

echo バックアップ完了。終了コード: %RC%

pause
