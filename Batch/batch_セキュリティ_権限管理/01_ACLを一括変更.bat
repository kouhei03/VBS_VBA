@echo off
rem 概要: 「ACLを一括変更」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "01_ACLを一括変更.bat"
rem 実行結果（想定結果）: 「ACLを一括変更」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
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
set "ACCOUNT="
set "RIGHT="
set /p "TARGET=アクセス権を変更するフォルダ: "
set /p "ACCOUNT=許可するユーザーまたはグループ（例: DOMAIN\User）: "
echo 1: 読み取り  2: 変更  3: フルコントロール
set /p "RIGHT=付与する権限 (1-3): "

if not exist "%TARGET%\." (
    echo フォルダが見つかりません。

    pause
    exit /b 1
)

set "BACKUP=%~dp0acl_backup.txt"

icacls "%TARGET%" /save "%BACKUP%" /T /C

if errorlevel 1 (
    echo ACL のバックアップに失敗しました。変更を中止します。

    pause
    exit /b 1
)

echo ACL のバックアップ: %BACKUP%
set "CONFIRM="
set /p "CONFIRM=対象フォルダ配下へ権限を適用する場合は APPLY と入力してください: "
if not "%CONFIRM%"=="APPLY" exit /b 0
if "%RIGHT%"=="1" icacls "%TARGET%" /grant "%ACCOUNT%:(OI)(CI)R" /T /C
if "%RIGHT%"=="2" icacls "%TARGET%" /grant "%ACCOUNT%:(OI)(CI)M" /T /C
if "%RIGHT%"=="3" icacls "%TARGET%" /grant "%ACCOUNT%:(OI)(CI)F" /T /C
if not "%RIGHT%"=="1" if not "%RIGHT%"=="2" if not "%RIGHT%"=="3" echo 権限の選択が不正です。

pause
