@echo off
rem 概要: 「Dockerイメージをログインしてpush」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "13_Dockerイメージをログインしてpush.bat"
rem 実行結果（想定結果）: 「Dockerイメージをログインしてpush」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: Dockerレジストリへログイン、Dockerイメージをレジストリへpush。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set "IMAGE="
set "REGISTRY="
set /p "REGISTRY=レジストリのホスト名（Docker Hub は空欄）: "
echo docker login を実行します。パスワードをこの BAT に記載しないでください。

if "%REGISTRY%"=="" (
    docker login
) else (
    docker login "%REGISTRY%"
)

if errorlevel 1 (
    echo ログインに失敗しました。

    pause
    exit /b 1
)

set /p "IMAGE=push する完全なイメージ名とタグ（例: user/app:1.0）: "
if "%IMAGE%"=="" exit /b 1

docker image inspect "%IMAGE%" >nul 2>&1

if errorlevel 1 (
    echo ローカルにそのイメージがありません。先にビルドしてください。

    pause
    exit /b 1
)

set "CONFIRM="
set /p "CONFIRM=指定イメージをレジストリへ push する場合は PUSH と入力: "
if not "%CONFIRM%"=="PUSH" exit /b 0

docker push "%IMAGE%"
if errorlevel 1 echo イメージの push に失敗しました。

pause
