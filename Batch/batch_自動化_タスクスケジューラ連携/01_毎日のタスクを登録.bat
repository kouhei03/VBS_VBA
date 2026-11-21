@echo off
rem 概要: 「毎日のタスクを登録」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "01_毎日のタスクを登録.bat"
rem 実行結果（想定結果）: 「毎日のタスクを登録」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: タスクスケジューラのタスクを照会、タスクスケジューラにタスクを登録。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set "TASKNAME="
set "TASKRUN="
set "STARTTIME="
set /p "TASKNAME=新しいタスク名（既存タスク名は指定しない）: "
if "%TASKNAME%"=="" exit /b 1

schtasks /query /tn "%TASKNAME%" >nul 2>&1

if not errorlevel 1 (
    echo 同名のタスクが既にあります。既存タスクは変更しません。

    pause
    exit /b 1
)

set /p "TASKRUN=実行するプログラムと引数: "
set /p "STARTTIME=毎日の開始時刻（24時間表記 HH:MM）: "
if "%TASKRUN%"=="" exit /b 1

schtasks /create /sc DAILY /tn "%TASKNAME%" /tr "%TASKRUN%" /st "%STARTTIME%"
if errorlevel 1 echo タスクを登録できませんでした。時刻や権限を確認してください。

pause
