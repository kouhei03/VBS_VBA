@echo off
rem 概要: 「AWSプロファイルを確認して切替」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "15_AWSプロファイルを確認して切替.bat"
rem 実行結果（想定結果）: 「AWSプロファイルを確認して切替」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: AWSプロファイルの認証先を確認。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
echo 利用可能な AWS CLI プロファイル:

aws configure list-profiles

if errorlevel 1 (
    echo プロファイルを取得できませんでした。

    pause
    exit /b 1
)

set "PROFILE="
set /p "PROFILE=使用するプロファイル名: "
if "%PROFILE%"=="" exit /b 1

aws sts get-caller-identity --profile "%PROFILE%"

if errorlevel 1 (
    echo 認証を確認できませんでした。プロファイルは変更されていません。

    pause
    exit /b 1
)

echo この BAT の実行中だけ AWS_PROFILE を設定します。設定ファイルは変更しません。
set "AWS_PROFILE=%PROFILE%"
echo 以降のAWSコマンドはこのウィンドウ内で選択したプロファイルを使用します。
cmd /k
