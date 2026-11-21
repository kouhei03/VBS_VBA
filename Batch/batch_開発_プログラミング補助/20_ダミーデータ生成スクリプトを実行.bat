@echo off
rem 概要: 「ダミーデータ生成スクリプトを実行」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "20_ダミーデータ生成スクリプトを実行.bat"
rem 実行結果（想定結果）: 「ダミーデータ生成スクリプトを実行」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では本文内のWindowsまたはPowerShellコマンドを実行します。分岐がある場合は選択条件に従います。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set "SCRIPT="
set "ARGS="
set /p "SCRIPT=実行するダミーデータ生成スクリプトのフルパス: "
set /p "ARGS=スクリプト引数（不要なら空欄）: "

if not exist "%SCRIPT%" (
    echo スクリプトが見つかりません。

    pause
    exit /b 1
)

echo 開発・テスト用データだけを生成することを確認してください。
set "CONFIRM="
set /p "CONFIRM=続行する場合は GENERATE と入力してください: "
if not "%CONFIRM%"=="GENERATE" exit /b 0
for %%F in ("%SCRIPT%") do set "EXT=%%~xF"

if /i "%EXT%"==".py" (
    python "%SCRIPT%" %ARGS%
) else if /i "%EXT%"==".ps1" (
    powershell -NoProfile -File "%SCRIPT%" %ARGS%
) else if /i "%EXT%"==".bat" (
    call "%SCRIPT%" %ARGS%
) else (
    echo 対応する拡張子は .py、.ps1、.bat です。
    exit /b 1
)

if errorlevel 1 echo データ生成スクリプトがエラー終了しました。

pause
