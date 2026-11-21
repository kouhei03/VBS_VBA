@echo off
rem 概要: 「ポート使用プロセスを確認して終了」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "16_ポート使用プロセスを確認して終了.bat"
rem 実行結果（想定結果）: 「ポート使用プロセスを確認して終了」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: 管理者権限の有無を確認、Get-NetTCPConnectionでTCP接続を取得、Get-Processで対象プロセスを取得、Stop-Processで指定プロセスを停止。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
net session >nul 2>&1

if errorlevel 1 (
    echo 管理者として実行してください。

    pause
    exit /b 1
)

set "PORT="
set /p "PORT=確認するローカル TCP ポート番号: "

powershell -NoProfile -Command "$p=0; if(-not [int]::TryParse($env:PORT,[ref]$p) -or $p -lt 1 -or $p -gt 65535){throw 'ポート番号は1～65535で指定してください'}; $items=@(Get-NetTCPConnection -LocalPort $p -State Listen -ErrorAction SilentlyContinue); if($items.Count -eq 0){Write-Host '待受プロセスは見つかりません。'; exit 0}; $items | ForEach-Object {Get-Process -Id $_.OwningProcess -ErrorAction SilentlyContinue | Select-Object Id,ProcessName,Path -Unique}; $answer=Read-Host '表示されたプロセスを強制終了する場合は KILL と入力'; if($answer -cne 'KILL'){Write-Host '中止しました。'; exit 0}; foreach($id in ($items.OwningProcess | Sort-Object -Unique)){if($id -le 4){throw 'システムプロセスは終了できません'}; Stop-Process -Id $id -Force -ErrorAction Stop}"
if errorlevel 1 echo プロセスを終了できませんでした。

pause
