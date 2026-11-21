@echo off
rem 概要: 「作業フォルダ指定で画像編集ソフト起動」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "13_作業フォルダ指定で画像編集ソフト起動.bat"
rem 実行結果（想定結果）: 「作業フォルダ指定で画像編集ソフト起動」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: Start-Processでアプリを起動。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。

chcp 65001 >nul
setlocal
set "APP="
set "WORKDIR="
set /p "APP=画像編集ソフトの exe フルパス: "
set /p "WORKDIR=作業用ディレクトリ: "

if not exist "%APP%" (
    echo アプリが見つかりません。

    pause
    exit /b 1
)

if not exist "%WORKDIR%\." (
    echo 作業用ディレクトリが見つかりません。

    pause
    exit /b 1
)

powershell -NoProfile -Command "Start-Process -FilePath $env:APP -WorkingDirectory $env:WORKDIR"
