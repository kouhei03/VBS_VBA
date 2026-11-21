@echo off
rem 概要: 「スタートアップ登録を管理」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "05_スタートアップ登録を管理.bat"
rem 実行結果（想定結果）: 「スタートアップ登録を管理」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: PowerShellのNew-Objectで対象を処理、Remove-Itemで対象を削除。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
echo 1: 現在のユーザーのスタートアップに BAT のショートカットを追加
echo 2: このスクリプトが作成したショートカットを削除
set "ACTION="
set /p "ACTION=操作を選択してください (1/2): "
if "%ACTION%"=="1" goto add_startup
if "%ACTION%"=="2" goto remove_startup
exit /b 1

:add_startup

set "TARGET="
set "LINKNAME="
set /p "TARGET=起動する BAT ファイルのフルパス: "

if not exist "%TARGET%" (
    echo ファイルが見つかりません。

    pause
    exit /b 1
)

set /p "LINKNAME=ショートカット名（拡張子なし）: "
if "%LINKNAME%"=="" exit /b 1

powershell -NoProfile -Command "$startup=[Environment]::GetFolderPath('Startup'); $link=Join-Path $startup ($env:LINKNAME+'.lnk'); if(Test-Path -LiteralPath $link){throw '同名のショートカットが既にあります'}; $shell=New-Object -ComObject WScript.Shell; $shortcut=$shell.CreateShortcut($link); $shortcut.TargetPath=$env:TARGET; $shortcut.WorkingDirectory=Split-Path -Parent $env:TARGET; $shortcut.Save(); Write-Host ('登録しました: '+$link)"
if errorlevel 1 echo スタートアップに登録できませんでした。

pause
exit /b

:remove_startup

set "LINKNAME="
set "CONFIRM="
set /p "LINKNAME=削除するショートカット名（拡張子なし）: "
set /p "CONFIRM=スタートアップ内の指定ショートカットを削除する場合は REMOVE と入力: "
if not "%CONFIRM%"=="REMOVE" exit /b 0

powershell -NoProfile -Command "$link=Join-Path ([Environment]::GetFolderPath('Startup')) ($env:LINKNAME+'.lnk'); if(-not (Test-Path -LiteralPath $link)){throw 'ショートカットが見つかりません'}; Remove-Item -LiteralPath $link -ErrorAction Stop"
if errorlevel 1 echo ショートカットを削除できませんでした。

pause
