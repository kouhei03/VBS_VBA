@echo off
rem 概要: 「テキストの行数をカウント」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "12_テキストの行数をカウント.bat"
rem 実行結果（想定結果）: 「テキストの行数をカウント」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では本文内のWindowsまたはPowerShellコマンドを実行します。分岐がある場合は選択条件に従います。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。

chcp 65001 >nul
setlocal
set "FILE="
set /p "FILE=行数を数えるテキストファイル: "

if not exist "%FILE%" (
    echo ファイルが見つかりません。

    pause
    exit /b 1
)

powershell -NoProfile -Command "$count=0; foreach($line in [IO.File]::ReadLines($env:FILE)){$count++}; Write-Host ('行数: '+$count)"

pause
