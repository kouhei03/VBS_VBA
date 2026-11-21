@echo off
rem 概要: 「フォルダを管理者専用にロック解除」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "02_フォルダを管理者専用にロック解除.bat"
rem 実行結果（想定結果）: 「フォルダを管理者専用にロック解除」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: 管理者権限の有無を確認。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
net session >nul 2>&1

if errorlevel 1 (
    echo 管理者として実行してください。

    pause
    exit /b 1
)

set "TARGET="
set "BACKUP="
echo ACL を保存してから、指定フォルダを Administrators と SYSTEM のみ利用可能にします。
echo 解除時は同じバックアップファイルを指定してください。
echo バックアップは対象フォルダの外に保存してください。
set /p "TARGET=対象フォルダ: "

if not exist "%TARGET%\." (
    echo フォルダが見つかりません。

    pause
    exit /b 1
)

echo 1: ロック解除（ACL をバックアップから復元）
echo 2: ロック（現在の ACL を保存して管理者専用にする）
set "ACTION="
set /p "ACTION=操作を選択してください (1/2): "
if "%ACTION%"=="1" goto unlock_folder
if "%ACTION%"=="2" goto lock_folder
exit /b 1

:unlock_folder

set /p "BACKUP=復元する ACL バックアップファイル: "

if not exist "%BACKUP%" (
    echo ACL バックアップが見つかりません。

    pause
    exit /b 1
)

icacls "%TARGET%" /restore "%BACKUP%" /C

pause
exit /b

:lock_folder

set /p "BACKUP=新規作成する ACL バックアップファイル: "

if exist "%BACKUP%" (
    echo バックアップファイルが既にあります。別の名前を指定してください。

    pause
    exit /b 1
)

set "CONFIRM="
set /p "CONFIRM=現在の ACL を上書きします。続行する場合は LOCK と入力: "
if not "%CONFIRM%"=="LOCK" exit /b 0

icacls "%TARGET%" /save "%BACKUP%" /T /C

if errorlevel 1 (
    echo ACL 保存に失敗したため、ロックしません。

    pause
    exit /b 1
)

icacls "%TARGET%" /reset /T /C
if errorlevel 1 goto restore_acl

icacls "%TARGET%" /inheritance:r /T /C
if errorlevel 1 goto restore_acl

icacls "%TARGET%" /grant:r "*S-1-5-18:(OI)(CI)F" "*S-1-5-32-544:(OI)(CI)F" /T /C
if errorlevel 1 goto restore_acl
echo ロック解除には、この BAT を管理者として起動し、保存した ACL バックアップを指定してください。

pause
exit /b

:restore_acl

echo ACL の変更に失敗しました。保存した ACL の復元を試みます。

icacls "%TARGET%" /restore "%BACKUP%" /C
echo 復元コマンドが失敗した場合は、バックアップファイルを使って手動で ACL を復元してください。

pause
exit /b 1
