@echo off
rem 概要: 「バックアップからファイルを復元」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "12_バックアップからファイルを復元.bat"
rem 実行結果（想定結果）: 「バックアップからファイルを復元」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: ファイルをコピー。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set /p "BACKUP_FILE=復元元ファイルのフルパス: "
set /p "DEST_FILE=復元先ファイルのフルパス: "

if not exist "%BACKUP_FILE%" (
    echo バックアップファイルが見つかりません。

    pause
    exit /b 1
)

if exist "%DEST_FILE%" (
    echo 復元先に同名ファイルがあります。上書きすると現在の内容は失われます。
    set /p "CONFIRM=上書きする場合は OVERWRITE と入力してください: "
    if not "%CONFIRM%"=="OVERWRITE" exit /b 0
)

for %%P in ("%DEST_FILE%") do if not exist "%%~dpP." mkdir "%%~dpP"
copy /Y "%BACKUP_FILE%" "%DEST_FILE%"
if errorlevel 1 echo 復元に失敗しました。

pause
