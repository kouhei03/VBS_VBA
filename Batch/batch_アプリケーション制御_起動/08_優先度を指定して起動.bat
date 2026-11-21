@echo off
rem 概要: 「優先度を指定して起動」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "08_優先度を指定して起動.bat"
rem 実行結果（想定結果）: 「優先度を指定して起動」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: Start-Processでアプリを起動。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set "APP="
set "ARGS="
set "PRIORITY="
set /p "APP=起動する exe のフルパス: "
set /p "ARGS=起動引数（不要なら空欄）: "

if not exist "%APP%" (
    echo 実行ファイルが見つかりません。

    pause
    exit /b 1
)

echo 1: High（高）
echo 2: AboveNormal（通常以上）
set /p "PRIORITY=優先度を選択してください (1/2): "

powershell -NoProfile -Command "$class=if($env:PRIORITY -eq '1'){'High'}elseif($env:PRIORITY -eq '2'){'AboveNormal'}else{throw '1 または 2 を選択してください'}; $p=Start-Process -FilePath $env:APP -ArgumentList $env:ARGS -PassThru; Start-Sleep -Milliseconds 500; $p.Refresh(); if(-not $p.HasExited){$p.PriorityClass=$class}"
if errorlevel 1 echo 優先度を設定できませんでした。Realtime はシステム応答に影響するため選択肢に含めていません。
