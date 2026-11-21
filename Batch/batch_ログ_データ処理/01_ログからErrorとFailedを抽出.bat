@echo off
rem 概要: 「ログからErrorとFailedを抽出」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "01_ログからErrorとFailedを抽出.bat"
rem 実行結果（想定結果）: 「ログからErrorとFailedを抽出」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: Select-Stringで指定キーワードを検索、Set-Contentで結果をファイルへ書き込み。複数の分岐がある場合は条件に応じて実行します。
rem   3. 結果やログを本文で指定したファイルへ保存し、必要に応じて画面にも表示します。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set "LOG="
set "OUTPUT="
set /p "LOG=検索するログファイル: "
set /p "OUTPUT=抽出結果の保存先: "

if not exist "%LOG%" (
    echo ログファイルが見つかりません。

    pause
    exit /b 1
)

if /i "%LOG%"=="%OUTPUT%" (
    echo 入力ログと出力先は別にしてください。
    exit /b 1
)

powershell -NoProfile -Command "$ErrorActionPreference='Stop'; Select-String -LiteralPath $env:LOG -Pattern 'Error|Failed' | ForEach-Object {$_.Line} | Set-Content -LiteralPath $env:OUTPUT -Encoding UTF8"
if errorlevel 1 echo 抽出に失敗しました。

pause
