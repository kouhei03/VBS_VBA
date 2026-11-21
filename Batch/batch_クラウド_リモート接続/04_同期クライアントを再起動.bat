@echo off
rem 概要: 「同期クライアントを再起動」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "04_同期クライアントを再起動.bat"
rem 実行結果（想定結果）: 「同期クライアントを再起動」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: Get-Processで対象プロセスを取得、Stop-Processで指定プロセスを停止、Start-Processでアプリを起動。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
echo 1: OneDrive
echo 2: Google Drive for desktop
set "APP="
set /p "APP=再起動するクライアント (1/2): "
if "%APP%"=="1" goto onedrive
if "%APP%"=="2" goto googledrive
exit /b 1

:onedrive

powershell -NoProfile -Command "$path=Join-Path $env:LOCALAPPDATA 'Microsoft\OneDrive\OneDrive.exe'; if(-not (Test-Path -LiteralPath $path)){throw 'OneDrive.exe が見つかりません'}; Get-Process OneDrive -ErrorAction SilentlyContinue | Stop-Process -Force; Start-Sleep -Seconds 2; Start-Process -FilePath $path"
goto result

:googledrive

powershell -NoProfile -Command "$paths=@((Join-Path $env:ProgramFiles 'Google\Drive File Stream\GoogleDriveFS.exe'),(Join-Path $env:LOCALAPPDATA 'Google\DriveFS\GoogleDriveFS.exe')); $path=$paths | Where-Object {Test-Path -LiteralPath $_} | Select-Object -First 1; if(-not $path){throw 'GoogleDriveFS.exe が見つかりません'}; Get-Process GoogleDriveFS -ErrorAction SilentlyContinue | Stop-Process -Force; Start-Sleep -Seconds 2; Start-Process -FilePath $path"

:result

if errorlevel 1 echo クライアントを再起動できませんでした。インストール場所を確認してください。

pause
