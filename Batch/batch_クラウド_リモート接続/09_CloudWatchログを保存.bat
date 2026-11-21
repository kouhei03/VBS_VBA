@echo off
rem 概要: 「CloudWatchログを保存」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "09_CloudWatchログを保存.bat"
rem 実行結果（想定結果）: 「CloudWatchログを保存」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: CloudWatch Logsからイベントを取得。複数の分岐がある場合は条件に応じて実行します。
rem   3. 結果やログを本文で指定したファイルへ保存し、必要に応じて画面にも表示します。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set "GROUP="
set "STREAM="
set "REGION="
set "OUTPUT=%~dp0cloudwatch_log.txt"
set /p "GROUP=CloudWatch ロググループ名: "
set /p "STREAM=ログストリーム名: "
set /p "REGION=AWS リージョン: "
if "%GROUP%"=="" exit /b 1
if "%STREAM%"=="" exit /b 1

aws logs get-log-events --log-group-name "%GROUP%" --log-stream-name "%STREAM%" --region "%REGION%" --start-from-head --output text > "%OUTPUT%"

if errorlevel 1 (
    echo ログを取得できませんでした。

    pause
    exit /b 1
)

echo ログを保存しました: %OUTPUT%

pause
