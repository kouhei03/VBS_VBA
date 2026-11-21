@echo off
rem 概要: 「今日のメモをカレンダーに追記」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "18_今日のメモをカレンダーに追記.bat"
rem 実行結果（想定結果）: 「今日のメモをカレンダーに追記」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: Add-Contentでログへ追記。複数の分岐がある場合は条件に応じて実行します。
rem   3. 結果やログを本文で指定したファイルへ保存し、必要に応じて画面にも表示します。

chcp 65001 >nul
setlocal
set "CALENDAR=%~dp0calendar.txt"
set "NOTE="
set /p "NOTE=今日のメモ: "
if "%NOTE%"=="" exit /b 1

powershell -NoProfile -Command "$line='[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd'),$env:NOTE; Add-Content -LiteralPath $env:CALENDAR -Value $line -Encoding UTF8"
echo 追記しました: %CALENDAR%
