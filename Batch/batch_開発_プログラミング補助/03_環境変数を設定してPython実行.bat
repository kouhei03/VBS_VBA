@echo off
rem 概要: 「環境変数を設定してPython実行」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "03_環境変数を設定してPython実行.bat"
rem 実行結果（想定結果）: 「環境変数を設定してPython実行」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では本文内のWindowsまたはPowerShellコマンドを実行します。分岐がある場合は選択条件に従います。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set "SCRIPT="
set "PYTHON=python"
set "VARNAME="
set "VARVALUE="
set /p "SCRIPT=実行する Python スクリプト: "
set /p "PYTHON=Python 実行ファイル（空欄なら python）: "
set /p "VARNAME=設定する環境変数名: "
set /p "VARVALUE=設定する環境変数値: "

if not exist "%SCRIPT%" (
    echo Python スクリプトが見つかりません。

    pause
    exit /b 1
)

if "%PYTHON%"=="" set "PYTHON=python"

powershell -NoProfile -Command "$name=$env:VARNAME; if($name -notmatch '^[A-Za-z_][A-Za-z0-9_]*$'){throw '環境変数名が不正です'}; [Environment]::SetEnvironmentVariable($name,$env:VARVALUE,'Process'); & $env:PYTHON $env:SCRIPT; exit $LASTEXITCODE"
if errorlevel 1 echo Python スクリプトがエラー終了しました。

pause
