@echo off
rem 概要: 「ローカルユーザーを追加削除」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "15_ローカルユーザーを追加削除.bat"
rem 実行結果（想定結果）: 「ローカルユーザーを追加削除」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: 管理者権限の有無を確認。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set "CURRENT_USER=%USERNAME%"
set "USERNAME="
net session >nul 2>&1

if errorlevel 1 (
    echo 管理者として実行してください。

    pause
    exit /b 1
)

echo 1: ローカルユーザーを追加
echo 2: ローカルユーザーを削除
set "ACTION="
set /p "ACTION=操作を選択してください (1/2): "
if "%ACTION%"=="1" goto add_user
if "%ACTION%"=="2" goto delete_user
echo 操作が選択されていません。
exit /b 1

:add_user

set /p "USERNAME=追加するユーザー名: "
if "%USERNAME%"=="" exit /b 1
echo パスワードを入力してください（入力文字は表示されません）。
net user "%USERNAME%" * /add
if errorlevel 1 echo ユーザーを追加できませんでした。

pause
exit /b

:delete_user

set /p "USERNAME=削除するユーザー名: "
if "%USERNAME%"=="" exit /b 1

if /i "%USERNAME%"=="%CURRENT_USER%" (
    echo 誤操作防止のため、このスクリプトからユーザーを削除できません。
    echo Windows のユーザー管理画面から対象を確認して操作してください。

    pause
    exit /b 1
)

echo 指定したアカウントを削除すると、そのユーザーのデータに影響する場合があります。
set "CONFIRM="
set /p "CONFIRM=削除する場合は DELETE と入力してください: "
if not "%CONFIRM%"=="DELETE" exit /b 0
net user "%USERNAME%" /delete
if errorlevel 1 echo ユーザーを削除できませんでした。

pause
