@echo off
rem 概要: 「日付連番バックアップフォルダ」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "04_日付連番バックアップフォルダ.bat"
rem 実行結果（想定結果）: 「日付連番バックアップフォルダ」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: Robocopyでファイルをコピーまたは同期。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set /p "SOURCE=バックアップ元フォルダ: "
set /p "PARENT=バックアップ先の親フォルダ: "

if not exist "%SOURCE%\." (
    echo バックアップ元が見つかりません。

    pause
    exit /b 1
)

if not exist "%PARENT%\." mkdir "%PARENT%"
for /f "usebackq delims=" %%D in (`powershell -NoProfile -Command "$date=Get-Date -Format yyyyMMdd; $n=1; do {$name='Backup_'+$date+'_'+('{0:D3}' -f $n); $path=Join-Path $env:PARENT $name; $n++} while(Test-Path -LiteralPath $path); Write-Output $path"`) do set "DEST=%%D"
mkdir "%DEST%"

robocopy "%SOURCE%" "%DEST%" /E /R:2 /W:2 /XJ
if errorlevel 8 echo バックアップに失敗しました。終了コード: %ERRORLEVEL%

pause
