@echo off
rem 概要: 「このBATを管理者として再起動」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "07_このBATを管理者として再起動.bat"
rem 実行結果（想定結果）: 「このBATを管理者として再起動」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. スクリプト内の設定値を使い、実行条件を確認します。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: Start-Processでアプリを起動。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
fltmc >nul 2>&1

if errorlevel 1 (
    set "BAT_PATH=%~f0"
    powershell -NoProfile -Command "Start-Process -FilePath $env:COMSPEC -ArgumentList ('/c ""'+$env:BAT_PATH+'""') -Verb RunAs"
    exit /b
)

echo 管理者権限で実行中です。

pause
