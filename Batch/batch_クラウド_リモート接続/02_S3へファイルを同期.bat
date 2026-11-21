@echo off
rem 概要: 「S3へファイルを同期」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "02_S3へファイルを同期.bat"
rem 実行結果（想定結果）: 「S3へファイルを同期」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: 指定フォルダをS3へ同期。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set "SOURCE="
set "DESTINATION="
set /p "SOURCE=同期元フォルダのフルパス: "

if not exist "%SOURCE%\." (
    echo 同期元フォルダが見つかりません。

    pause
    exit /b 1
)

set /p "DESTINATION=同期先 S3 URI（例: s3://bucket-name/path/）: "

powershell -NoProfile -Command "if($env:DESTINATION -notmatch '^s3://[A-Za-z0-9.-]+(?:/.*)?$'){throw 'S3 URI の形式が不正です'}"
if errorlevel 1 exit /b 1
echo --delete は指定しません。同期先に余分なファイルが残る場合があります。
set "CONFIRM="
set /p "CONFIRM=指定先へ同期する場合は SYNC と入力: "
if not "%CONFIRM%"=="SYNC" exit /b 0

aws s3 sync "%SOURCE%" "%DESTINATION%"
if errorlevel 1 echo 同期に失敗しました。AWS CLI の認証、権限、パスを確認してください。

pause
