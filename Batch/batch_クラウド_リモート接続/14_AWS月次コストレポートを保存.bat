@echo off
rem 概要: 「AWS月次コストレポートを保存」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "14_AWS月次コストレポートを保存.bat"
rem 実行結果（想定結果）: 「AWS月次コストレポートを保存」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: Cost Explorerから利用料金を取得。複数の分岐がある場合は条件に応じて実行します。
rem   3. 結果やログを本文で指定したファイルへ保存し、必要に応じて画面にも表示します。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set "REGION="
set "OUTPUT=%~dp0aws_monthly_cost.json"
set /p "REGION=AWS CLI リージョン（Cost Explorer API 利用リージョン）: "
for /f "delims=" %%D in ('powershell -NoProfile -Command "Get-Date ([datetime]::Today.AddDays(1-[datetime]::Today.Day).AddMonths(-1)) -Format yyyy-MM-dd"') do set "STARTDATE=%%D"
for /f "delims=" %%D in ('powershell -NoProfile -Command "Get-Date -Day 1 -Format yyyy-MM-dd"') do set "ENDDATE=%%D"

aws ce get-cost-and-usage --time-period Start=%STARTDATE%,End=%ENDDATE% --granularity MONTHLY --metrics UnblendedCost --region "%REGION%" --output json > "%OUTPUT%"

if errorlevel 1 (
    echo コスト情報を取得できませんでした。Cost Explorer の権限と有効化状態を確認してください。

    pause
    exit /b 1
)

echo 先月分のコストレポートを保存しました: %OUTPUT%

pause
