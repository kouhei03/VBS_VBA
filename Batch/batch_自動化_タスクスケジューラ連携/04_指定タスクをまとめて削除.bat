@echo off
rem 概要: 「指定タスクをまとめて削除」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "04_指定タスクをまとめて削除.bat"
rem 実行結果（想定結果）: 「指定タスクをまとめて削除」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: 指定タスクを削除。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
echo 削除するタスク名を ^| で区切って入力します（例: \TaskA^|\Folder\TaskB）。
set "TASKS="
set /p "TASKS=削除対象タスク: "
if "%TASKS%"=="" exit /b 1

powershell -NoProfile -Command "$env:TASKS -split '\|' | ForEach-Object {$_.Trim()} | Where-Object {$_} | ForEach-Object {Write-Host ('対象: '+$_)}"
set "CONFIRM="
set /p "CONFIRM=上記タスクを削除する場合は DELETE と入力してください: "
if not "%CONFIRM%"=="DELETE" exit /b 0

powershell -NoProfile -Command "$tasks=($env:TASKS -split '\|' | ForEach-Object {$_.Trim()} | Where-Object {$_}); $failed=$false; foreach($task in $tasks){Write-Host ('削除: '+$task); & schtasks.exe /delete /tn $task /f; if($LASTEXITCODE -ne 0){Write-Error ('削除失敗: '+$task); $failed=$true}}; if($failed){exit 1}"
if errorlevel 1 echo 一部またはすべてのタスクを削除できませんでした。

pause
