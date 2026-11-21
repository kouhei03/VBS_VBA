@echo off
rem 概要: 「プロキシ設定を切替」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "13_プロキシ設定を切替.bat"
rem 実行結果（想定結果）: 「プロキシ設定を切替」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: Set-ItemPropertyでレジストリ値を変更。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。

chcp 65001 >nul
setlocal
echo 1: プロキシを有効化
echo 2: プロキシを無効化
set "ACTION="
set /p "ACTION=操作を選択してください (1/2): "
if "%ACTION%"=="1" goto enable_proxy
if "%ACTION%"=="2" goto disable_proxy
echo 操作が選択されていません。
exit /b 1

:enable_proxy

set "PROXY_SERVER="
set /p "PROXY_SERVER=プロキシのアドレスとポート（例: proxy.example.com:8080）: "
if "%PROXY_SERVER%"=="" exit /b 1

powershell -NoProfile -Command "$p='HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings'; Set-ItemProperty -Path $p -Name ProxyServer -Value $env:PROXY_SERVER; Set-ItemProperty -Path $p -Name ProxyEnable -Value 1"
goto done

:disable_proxy

powershell -NoProfile -Command "Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings' -Name ProxyEnable -Value 0"

:done

echo 設定を変更しました。実行中のアプリは再起動が必要な場合があります。

pause
