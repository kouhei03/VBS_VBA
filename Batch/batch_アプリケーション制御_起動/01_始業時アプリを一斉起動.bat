@echo off
rem 概要: 「始業時アプリを一斉起動」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "01_始業時アプリを一斉起動.bat"
rem 実行結果（想定結果）: 「始業時アプリを一斉起動」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では本文内のWindowsまたはPowerShellコマンドを実行します。分岐がある場合は選択条件に従います。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。

chcp 65001 >nul
setlocal
set "BROWSER="
set "EXCEL="
set "TEAMS="
set /p "BROWSER=起動するブラウザーの exe パス（空欄で省略）: "
set /p "EXCEL=Excel の exe パス（空欄で省略）: "
set /p "TEAMS=Teams の exe パス（空欄で省略）: "
if not "%BROWSER%"=="" if exist "%BROWSER%" start "" "%BROWSER%"
if not "%EXCEL%"=="" if exist "%EXCEL%" start "" "%EXCEL%"
if not "%TEAMS%"=="" if exist "%TEAMS%" start "" "%TEAMS%"
echo 起動要求を送信しました。パスを確認してください。

pause
