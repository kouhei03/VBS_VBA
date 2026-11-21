@echo off
rem 概要: 「APIのJSONレスポンスを保存」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "20_APIのJSONレスポンスを保存.bat"
rem 実行結果（想定結果）: 「APIのJSONレスポンスを保存」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: Invoke-WebRequestでHTTP応答を取得、PowerShellのConvertFrom-Jsonで対象を処理。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set "URL="
set "OUTPUT="
set /p "URL=JSON API の HTTPS URL: "
set /p "OUTPUT=レスポンス保存先ファイル: "
if not exist "%OUTPUT%" goto save_response
set "CONFIRM="
set /p "CONFIRM=既存ファイルを上書きする場合は OVERWRITE と入力: "
if not "%CONFIRM%"=="OVERWRITE" exit /b 0

:save_response

powershell -NoProfile -Command "$ErrorActionPreference='Stop'; if($env:URL -notmatch '^https://'){throw 'https:// の URL を指定してください'}; $r=Invoke-WebRequest -Uri $env:URL -TimeoutSec 30 -UseBasicParsing; $r.Content | ConvertFrom-Json | Out-Null; [IO.File]::WriteAllText($env:OUTPUT,$r.Content,[Text.UTF8Encoding]::new($false)); Write-Host ('保存しました: '+$env:OUTPUT)"
if errorlevel 1 echo API の呼び出しまたは JSON 保存に失敗しました。

pause
