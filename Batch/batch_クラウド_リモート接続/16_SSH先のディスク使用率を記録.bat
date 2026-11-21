@echo off
rem 概要: 「SSH先のディスク使用率を記録」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "16_SSH先のディスク使用率を記録.bat"
rem 実行結果（想定結果）: 「SSH先のディスク使用率を記録」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: OpenSSHでリモート接続またはコマンド実行、Add-Contentでログへ追記。複数の分岐がある場合は条件に応じて実行します。
rem   3. 結果やログを本文で指定したファイルへ保存し、必要に応じて画面にも表示します。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set "REMOTE="
set "THRESHOLD="
set "LOG=%~dp0remote_disk_alert.log"
set /p "REMOTE=SSH 接続先（例: user@server.example.com）: "
set /p "THRESHOLD=警告する使用率（1-99%%）: "

powershell -NoProfile -Command "if($env:REMOTE -notmatch '^[A-Za-z0-9._-]+@[A-Za-z0-9.-]+$' -or $env:THRESHOLD -notmatch '^\d+$' -or [int]$env:THRESHOLD -lt 1 -or [int]$env:THRESHOLD -gt 99){throw '接続先または閾値が不正です'}"
if errorlevel 1 exit /b 1

powershell -NoProfile -Command "$lines=& ssh.exe -o BatchMode=yes $env:REMOTE 'df -P /'; if($LASTEXITCODE -ne 0){exit $LASTEXITCODE}; $row=$lines | Select-Object -Last 1; if($row -notmatch '\s(\d+)%\s+/'){throw 'df の出力を解析できません'}; $used=[int]$matches[1]; $line='{0} {1} usage={2}%%' -f (Get-Date -Format s),$env:REMOTE,$used; Add-Content -LiteralPath $env:LOG -Value $line -Encoding UTF8; Write-Host $line; if($used -ge [int]$env:THRESHOLD){Write-Warning 'ディスク使用率が閾値以上です'}"
if errorlevel 1 echo SSH 接続または容量確認に失敗しました。

pause
