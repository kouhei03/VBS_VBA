@echo off
rem 概要: 「指定プロセスを一括強制終了」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "02_指定プロセスを一括強制終了.bat"
rem 実行結果（想定結果）: 「指定プロセスを一括強制終了」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: Get-Processで対象プロセスを取得、Stop-Processで指定プロセスを停止。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set "PROCESSES="
set /p "PROCESSES=終了するプロセス名をカンマ区切りで入力（.exe なし）: "
if "%PROCESSES%"=="" exit /b 1
echo 未保存データが失われる可能性があります。

powershell -NoProfile -Command "$names=@($env:PROCESSES -split ',' | ForEach-Object {$_.Trim()} | Where-Object {$_}); $items=@(foreach($n in $names){Get-Process -Name $n -ErrorAction SilentlyContinue}); $items | Select-Object Id,ProcessName,Path -Unique; if($items.Count -eq 0){Write-Host '対象プロセスはありません。'; exit 0}; $answer=Read-Host '上記プロセスを強制終了する場合は KILL と入力'; if($answer -cne 'KILL'){Write-Host '中止しました。'; exit 0}; foreach($p in $items){Stop-Process -Id $p.Id -Force -ErrorAction Stop}"
if errorlevel 1 echo プロセスを終了できませんでした。

pause
