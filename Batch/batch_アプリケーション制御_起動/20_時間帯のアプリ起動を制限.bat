@echo off
rem 概要: 「時間帯のアプリ起動を制限」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "20_時間帯のアプリ起動を制限.bat"
rem 実行結果（想定結果）: 「時間帯のアプリ起動を制限」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: Get-Processで対象プロセスを取得、Stop-Processで指定プロセスを停止。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set "PROCESS="
set "START_TIME="
set "END_TIME="
set /p "PROCESS=制限するプロセス名（.exe なし）: "
set /p "START_TIME=制限開始時刻（24時間 HH:mm）: "
set /p "END_TIME=制限終了時刻（24時間 HH:mm）: "
if "%PROCESS%"=="" exit /b 1
echo このウィンドウが開いている間だけ監視します。停止は Ctrl+C です。
echo 制限時間中に対象アプリが起動すると強制終了します（未保存データに注意）。
set "CONFIRM="
set /p "CONFIRM=監視を開始する場合は BLOCK と入力してください: "
if not "%CONFIRM%"=="BLOCK" exit /b 0

powershell -NoProfile -Command "$start=[TimeSpan]::Zero; $end=[TimeSpan]::Zero; if(-not [TimeSpan]::TryParse($env:START_TIME,[ref]$start) -or -not [TimeSpan]::TryParse($env:END_TIME,[ref]$end)){throw '時刻は HH:mm 形式で指定してください'}; while($true){$now=(Get-Date).TimeOfDay; $blocked=if($start -le $end){$now -ge $start -and $now -lt $end}else{$now -ge $start -or $now -lt $end}; if($blocked){Get-Process -Name $env:PROCESS -ErrorAction SilentlyContinue | ForEach-Object {Write-Host ('制限時間内のため終了: '+$_.ProcessName); Stop-Process -Id $_.Id -Force -ErrorAction Continue}}; Start-Sleep -Seconds 5}"
