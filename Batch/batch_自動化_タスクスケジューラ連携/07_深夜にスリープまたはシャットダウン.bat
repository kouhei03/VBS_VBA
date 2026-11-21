@echo off
rem 概要: 「深夜にスリープまたはシャットダウン」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "07_深夜にスリープまたはシャットダウン.bat"
rem 実行結果（想定結果）: 「深夜にスリープまたはシャットダウン」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: Windowsをシャットダウンまたは再起動、タスクスケジューラのタスクを照会、タスクスケジューラにタスクを登録。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
echo 注意: 登録後は毎日 00:00 に指定処理が実行されます。
echo 1: 60秒後にシャットダウン
echo 2: スリープ（PCの電源設定により休止状態になる場合があります）
set "ACTION="
set /p "ACTION=操作を選択してください (1/2): "
if "%ACTION%"=="1" set "TASKRUN=shutdown.exe /s /t 60"
if "%ACTION%"=="2" set "TASKRUN=rundll32.exe powrprof.dll,SetSuspendState 0,1,0"
if not defined TASKRUN exit /b 1
set "TASKNAME="
set /p "TASKNAME=新しいタスク名: "
if "%TASKNAME%"=="" exit /b 1

schtasks /query /tn "%TASKNAME%" >nul 2>&1

if not errorlevel 1 (
    echo 同名のタスクが既にあります。既存タスクは変更しません。

    pause
    exit /b 1
)

set "CONFIRM="
set /p "CONFIRM=毎日深夜に実行するタスクを登録する場合は SCHEDULE と入力: "
if not "%CONFIRM%"=="SCHEDULE" exit /b 0

schtasks /create /sc DAILY /st 00:00 /tn "%TASKNAME%" /tr "%TASKRUN%"
if errorlevel 1 echo タスクを登録できませんでした。

pause
