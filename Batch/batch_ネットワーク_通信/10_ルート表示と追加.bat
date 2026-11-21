@echo off
rem 概要: 「ルート表示と追加」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "10_ルート表示と追加.bat"
rem 実行結果（想定結果）: 「ルート表示と追加」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: 管理者権限の有無を確認。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
echo 現在のルーティングテーブル:
route print
echo.
echo 1: 表示のみ
echo 2: 経路を追加（管理者権限が必要）
set "ACTION="
set /p "ACTION=操作を選択してください (1/2): "
if "%ACTION%"=="1" exit /b 0
if not "%ACTION%"=="2" exit /b 1
net session >nul 2>&1

if errorlevel 1 (
    echo 管理者として実行してください。

    pause
    exit /b 1
)

set "NETWORK="
set "MASK="
set "GATEWAY="
set "METRIC="
set /p "NETWORK=宛先ネットワーク（例: 10.0.0.0）: "
set /p "MASK=サブネットマスク: "
set /p "GATEWAY=ゲートウェイ: "
set /p "METRIC=メトリック（例: 10）: "
echo.
echo 入力した宛先・マスク・ゲートウェイ・メトリックの経路を追加します。
set "CONFIRM="
set /p "CONFIRM=続行する場合は ADD と入力してください: "
if not "%CONFIRM%"=="ADD" exit /b 0

powershell -NoProfile -Command "$ips=@($env:NETWORK,$env:MASK,$env:GATEWAY); foreach($ip in $ips){$parsed=$null; if(-not [Net.IPAddress]::TryParse($ip,[ref]$parsed) -or $parsed.AddressFamily -ne [Net.Sockets.AddressFamily]::InterNetwork){throw 'IPv4 アドレスを確認してください'}}; $m=0; if(-not [int]::TryParse($env:METRIC,[ref]$m) -or $m -lt 1){throw 'メトリックは1以上の整数にしてください'}; & route.exe ADD $env:NETWORK MASK $env:MASK $env:GATEWAY METRIC $m; if($LASTEXITCODE -ne 0){exit $LASTEXITCODE}"
if errorlevel 1 echo 経路を追加できませんでした。

pause
