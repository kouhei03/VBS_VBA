@echo off
rem 概要: 「WSLまたはDockerを起動停止」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "11_WSLまたはDockerを起動停止.bat"
rem 実行結果（想定結果）: 「WSLまたはDockerを起動停止」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: Docker CLIでイメージまたはコンテナを操作、Windowsをシャットダウンまたは再起動。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。

chcp 65001 >nul
setlocal
echo 1: WSL を起動
echo 2: WSL をすべて停止
echo 3: Docker Desktop を起動
echo 4: Docker Desktop を停止（docker desktop stop が必要）
set "ACTION="
set /p "ACTION=操作を選択してください (1-4): "
if "%ACTION%"=="1" goto start_wsl
if "%ACTION%"=="2" goto stop_wsl
if "%ACTION%"=="3" goto start_docker
if "%ACTION%"=="4" goto stop_docker
echo 操作が選択されていません。
exit /b 1

:start_wsl

start "" wsl.exe
goto done

:stop_wsl

set "CONFIRM="
set /p "CONFIRM=全 WSL ディストリビューションを停止する場合は STOP と入力: "
if /i not "%CONFIRM%"=="STOP" exit /b 0
wsl.exe --shutdown
goto done

:start_docker

if not exist "%ProgramFiles%\Docker\Docker\Docker Desktop.exe" (
    echo Docker Desktop が標準パスにありません。
    exit /b 1
)

start "" "%ProgramFiles%\Docker\Docker\Docker Desktop.exe"
goto done

:stop_docker

docker desktop stop

:done

pause
