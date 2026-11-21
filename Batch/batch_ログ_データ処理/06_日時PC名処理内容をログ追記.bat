@echo off
rem 概要: 「日時PC名処理内容をログ追記」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "06_日時PC名処理内容をログ追記.bat"
rem 実行結果（想定結果）: 「日時PC名処理内容をログ追記」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: Add-Contentでログへ追記。複数の分岐がある場合は条件に応じて実行します。
rem   3. 結果やログを本文で指定したファイルへ保存し、必要に応じて画面にも表示します。

chcp 65001 >nul
setlocal
set "LOG=%~dp0activity.log"
set "ACTIVITY="
set /p "ACTIVITY=記録する処理内容: "
if "%ACTIVITY%"=="" exit /b 1

powershell -NoProfile -Command "$line='{0}{1}{2}{1}{3}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),[char]9,$env:COMPUTERNAME,$env:ACTIVITY; Add-Content -LiteralPath $env:LOG -Value $line -Encoding UTF8"
echo 記録しました: %LOG%
