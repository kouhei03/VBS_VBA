@echo off
rem 概要: 「EC2へSSH接続」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "03_EC2へSSH接続.bat"
rem 実行結果（想定結果）: 「EC2へSSH接続」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: OpenSSHでリモート接続またはコマンド実行。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set "USERHOST="
set "PORT=22"
set /p "USERHOST=接続先（例: ec2-user@ec2-host.example.com）: "
if "%USERHOST%"=="" exit /b 1
set /p "PORT=SSH ポート（通常は 22）: "

powershell -NoProfile -Command "if($env:USERHOST -notmatch '^[A-Za-z0-9._-]+@[A-Za-z0-9.-]+$' -or $env:PORT -notmatch '^\d{1,5}$' -or [int]$env:PORT -lt 1 -or [int]$env:PORT -gt 65535){throw '接続先またはポートが不正です'}"
if errorlevel 1 exit /b 1
echo ローカルの SSH 鍵または SSH agent を使用します。秘密鍵をこの BAT に保存しないでください。

ssh -p %PORT% "%USERHOST%"
if errorlevel 1 echo SSH 接続に失敗しました。接続先、鍵、ネットワークを確認してください。

pause
