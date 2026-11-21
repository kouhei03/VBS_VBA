@echo off
rem 概要: 「HTMLをダウンロード」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "07_HTMLをダウンロード.bat"
rem 実行結果（想定結果）: 「HTMLをダウンロード」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: HTTPリクエストまたはファイル転送を実行。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
where curl.exe >nul 2>&1

if errorlevel 1 (
    echo curl.exe が見つかりません。

    pause
    exit /b 1
)

set "URL="
set "OUTPUT="
set /p "URL=取得する HTML の URL: "
set /p "OUTPUT=保存するファイルパス（例: page.html）: "

if exist "%OUTPUT%" (
    echo 保存先ファイルが既にあります。別の名前を指定してください。

    pause
    exit /b 1
)

curl.exe -L --fail --max-time 60 "%URL%" -o "%OUTPUT%"
if errorlevel 1 echo ダウンロードに失敗しました。

pause
