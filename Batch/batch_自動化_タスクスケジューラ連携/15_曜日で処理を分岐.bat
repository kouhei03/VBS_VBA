@echo off
rem 概要: 「曜日で処理を分岐」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "15_曜日で処理を分岐.bat"
rem 実行結果（想定結果）: 「曜日で処理を分岐」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. スクリプト内の設定値を使い、実行条件を確認します。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では本文内のWindowsまたはPowerShellコマンドを実行します。分岐がある場合は選択条件に従います。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。

chcp 65001 >nul
setlocal
for /f "delims=" %%D in ('powershell -NoProfile -Command "(Get-Date).DayOfWeek.ToString()"') do set "WEEKDAY=%%D"
echo 今日の曜日: %WEEKDAY%
if /i "%WEEKDAY%"=="Monday" goto monday
if /i "%WEEKDAY%"=="Tuesday" goto tuesday
if /i "%WEEKDAY%"=="Wednesday" goto wednesday
if /i "%WEEKDAY%"=="Thursday" goto thursday
if /i "%WEEKDAY%"=="Friday" goto friday
if /i "%WEEKDAY%"=="Saturday" goto saturday
if /i "%WEEKDAY%"=="Sunday" goto sunday
echo 曜日を取得できませんでした。
exit /b 1

:monday

echo 月曜日の処理をここに記述してください。
goto finish

:tuesday

echo 火曜日の処理をここに記述してください。
goto finish

:wednesday

echo 水曜日の処理をここに記述してください。
goto finish

:thursday

echo 木曜日の処理をここに記述してください。
goto finish

:friday

echo 金曜日の処理をここに記述してください。
goto finish

:saturday

echo 土曜日の処理をここに記述してください。
goto finish

:sunday

echo 日曜日の処理をここに記述してください。

:finish
