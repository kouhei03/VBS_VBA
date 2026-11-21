@echo off
rem 概要: 「ローカルDBをダンプ」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "10_ローカルDBをダンプ.bat"
rem 実行結果（想定結果）: 「ローカルDBをダンプ」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では本文内のWindowsまたはPowerShellコマンドを実行します。分岐がある場合は選択条件に従います。
rem   3. 結果やログを本文で指定したファイルへ保存し、必要に応じて画面にも表示します。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
echo 1: MySQL / MariaDB（mysqldump）
echo 2: PostgreSQL（pg_dump）
set "TYPE="
set /p "TYPE=データベースを選択してください (1/2): "
set "HOST=localhost"
set "DATABASE="
set "USER="
set "OUTDIR="
set /p "HOST=DB ホスト（既定 localhost）: "
set /p "DATABASE=データベース名: "
set /p "USER=DB ユーザー名: "
set /p "OUTDIR=ダンプ保存フォルダ: "
if "%HOST%"=="" set "HOST=localhost"
if not exist "%OUTDIR%\." mkdir "%OUTDIR%"
for /f %%T in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd_HHmmss"') do set "STAMP=%%T"

if "%TYPE%"=="1" (
    mysqldump -h "%HOST%" -u "%USER%" -p "%DATABASE%" > "%OUTDIR%\%DATABASE%_%STAMP%.sql"
) else if "%TYPE%"=="2" (
    pg_dump -h "%HOST%" -U "%USER%" -W -F c -f "%OUTDIR%\%DATABASE%_%STAMP%.dump" "%DATABASE%"
) else (
    echo データベースが選択されていません。
    exit /b 1
)

if errorlevel 1 echo ダンプに失敗しました。資格情報とデータベース名を確認してください。

pause
