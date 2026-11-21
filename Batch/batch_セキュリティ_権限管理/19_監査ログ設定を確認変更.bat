@echo off
rem 概要: 「監査ログ設定を確認変更」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "19_監査ログ設定を確認変更.bat"
rem 実行結果（想定結果）: 「監査ログ設定を確認変更」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: 管理者権限の有無を確認、Windowsイベントログを参照または設定、Windows監査ポリシーを参照または設定。複数の分岐がある場合は条件に応じて実行します。
rem   3. 結果やログを本文で指定したファイルへ保存し、必要に応じて画面にも表示します。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
net session >nul 2>&1

if errorlevel 1 (
    echo 管理者として実行してください。

    pause
    exit /b 1
)

set "OUTPUT=%~dp0audit_log_settings.txt"
(
    echo ==== Security Event Log ====
    wevtutil gl Security
    echo ==== Audit Policy ====
    auditpol /get /category:*
) > "%OUTPUT%"
type "%OUTPUT%"
echo 1: 確認のみ
echo 2: Security ログ最大サイズを変更
echo 3: Security ログの保持方式を変更
set "ACTION="
set /p "ACTION=操作を選択してください (1/2/3): "
if "%ACTION%"=="1" exit /b 0
if "%ACTION%"=="2" goto set_size
if "%ACTION%"=="3" goto set_retention
exit /b 1

:set_size

set "SIZE_MB="
set /p "SIZE_MB=新しい最大サイズ（MB、現在値以上を推奨）: "
set "CONFIRM="
set /p "CONFIRM=ログ最大サイズを変更する場合は APPLY と入力してください: "
if not "%CONFIRM%"=="APPLY" exit /b 0

powershell -NoProfile -Command "$n=0; if(-not [int]::TryParse($env:SIZE_MB,[ref]$n) -or $n -lt 16){throw 'サイズは16 MB以上で指定してください'}; & wevtutil.exe sl Security ('/ms:'+$n*1MB); if($LASTEXITCODE -ne 0){exit $LASTEXITCODE}"
if errorlevel 1 echo ログサイズを変更できませんでした。

pause
exit /b

:set_retention

echo 1: 古いイベントを必要に応じて上書き
echo 2: 上書きせず保持し、満杯時に自動アーカイブ
set "RETENTION="
set /p "RETENTION=保持方式を選択してください (1/2): "
if "%RETENTION%"=="1" set "RETENTION_ARGS=/rt:false"
if "%RETENTION%"=="2" set "RETENTION_ARGS=/rt:true /ab:true"
if not defined RETENTION_ARGS exit /b 1
set "CONFIRM="
set /p "CONFIRM=Security ログの保持方式を変更する場合は APPLY と入力してください: "
if not "%CONFIRM%"=="APPLY" exit /b 0

wevtutil sl Security %RETENTION_ARGS%
if errorlevel 1 echo 保持方式を変更できませんでした。

pause
