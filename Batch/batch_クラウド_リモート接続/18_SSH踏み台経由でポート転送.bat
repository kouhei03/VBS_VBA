@echo off
rem 概要: 「SSH踏み台経由でポート転送」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "18_SSH踏み台経由でポート転送.bat"
rem 実行結果（想定結果）: 「SSH踏み台経由でポート転送」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: OpenSSHでリモート接続またはコマンド実行、SSHでローカルポート転送トンネルを作成。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set "BASTION="
set "LOCALPORT="
set "REMOTEPORT="
set /p "BASTION=踏み台 SSH 接続先（例: user@bastion.example.com）: "
set /p "LOCALPORT=ローカル待受ポート: "
set /p "REMOTEPORT=踏み台から接続する宛先ポート: "

powershell -NoProfile -Command "if($env:BASTION -notmatch '^[A-Za-z0-9._-]+@[A-Za-z0-9.-]+$' -or $env:LOCALPORT -notmatch '^\d{1,5}$' -or $env:REMOTEPORT -notmatch '^\d{1,5}$' -or [int]$env:LOCALPORT -lt 1 -or [int]$env:LOCALPORT -gt 65535 -or [int]$env:REMOTEPORT -lt 1 -or [int]$env:REMOTEPORT -gt 65535){throw '接続先またはポートが不正です'}"
if errorlevel 1 exit /b 1
echo 接続先は踏み台上の 127.0.0.1:%REMOTEPORT% です。停止は Ctrl+C です。

ssh -N -L %LOCALPORT%:127.0.0.1:%REMOTEPORT% "%BASTION%"
if errorlevel 1 echo トンネル接続に失敗しました。
