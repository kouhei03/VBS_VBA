@echo off
rem 概要: 「USBストレージ制限を切替」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "08_USBストレージ制限を切替.bat"
rem 実行結果（想定結果）: 「USBストレージ制限を切替」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: 管理者権限の有無を確認、レジストリを参照または変更。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
net session >nul 2>&1

if errorlevel 1 (
    echo 管理者として実行してください。

    pause
    exit /b 1
)

echo 1: リムーバブル記憶域へのアクセスを禁止
echo 2: 禁止設定を解除
set "ACTION="
set /p "ACTION=操作を選択してください (1/2): "
if "%ACTION%"=="1" goto deny_usb
if "%ACTION%"=="2" goto allow_usb
exit /b 1

:deny_usb

echo USB ストレージの読み書きが全ユーザーに制限されます。入力機器には影響しません。
set "CONFIRM="
set /p "CONFIRM=続行する場合は DENY と入力してください: "
if not "%CONFIRM%"=="DENY" exit /b 0
reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\RemovableStorageDevices" /v Deny_All /t REG_DWORD /d 1 /f
goto update_policy

:allow_usb

reg delete "HKLM\SOFTWARE\Policies\Microsoft\Windows\RemovableStorageDevices" /v Deny_All /f

:update_policy

gpupdate /target:computer /force

pause
