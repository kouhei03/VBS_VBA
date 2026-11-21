@echo off
rem 概要: 「指定サービスを停止して無効化」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "06_指定サービスを停止して無効化.bat"
rem 実行結果（想定結果）: 「指定サービスを停止して無効化」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: 管理者権限の有無を確認、PowerShellのGet-Serviceで対象を処理、PowerShellのStop-Serviceで対象を処理、PowerShellのSet-Serviceで対象を処理。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
net session >nul 2>&1

if errorlevel 1 (
    echo 管理者として実行してください。

    pause
    exit /b 1
)

set "SERVICES="
set /p "SERVICES=無効化するサービス名をカンマ区切りで入力: "
if "%SERVICES%"=="" exit /b 1
echo 入力したサービスの状態を確認してから無効化します。

powershell -NoProfile -Command "$names=@($env:SERVICES -split ',' | ForEach-Object {$_.Trim()} | Where-Object {$_}); $blocked=@('RpcSs','DcomLaunch','PlugPlay','EventLog','MpsSvc','BFE','WinDefend','Power'); $bad=@($names | Where-Object {$blocked -contains $_}); if($bad.Count){Write-Error ('重要サービスは対象外です: '+($bad -join ', ')); exit 1}; Get-Service -Name $names -ErrorAction Stop | Format-Table Name,DisplayName,Status -AutoSize; $answer=Read-Host '上記サービスを停止し、スタートアップを無効にする場合は DISABLE と入力'; if($answer -cne 'DISABLE'){Write-Host '中止しました。'; exit 0}; foreach($name in $names){$svc=Get-Service -Name $name -ErrorAction Stop; if($svc.Status -eq 'Running'){Stop-Service -Name $name -ErrorAction Stop}; Set-Service -Name $name -StartupType Disabled -ErrorAction Stop}"
if errorlevel 1 echo サービスを変更できませんでした。対象名や依存関係を確認してください。

pause
