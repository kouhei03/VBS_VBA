@echo off
rem 概要: 「FTPSで一括転送」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "19_FTPSで一括転送.bat"
rem 実行結果（想定結果）: 「FTPSで一括転送」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: HTTPリクエストまたはファイル転送を実行、FTPによるファイル転送を実行。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
where curl.exe >nul 2>&1

if errorlevel 1 (
    echo curl.exe が見つかりません。

    pause
    exit /b 1
)

echo FTP は平文のため使わず、TLS 対応の FTPS 接続を使用します。
echo 1: フォルダ直下のファイルをアップロード
echo 2: FTPS のワイルドカード URL からダウンロード
set "ACTION="
set /p "ACTION=操作を選択してください (1/2): "
set "FTP_USER="
set "FTP_URL="
set /p "FTP_USER=FTPS ユーザー名（パスワードは curl が要求した時に入力）: "
set /p "FTP_URL=FTPS URL（例: ftps://server/path/）: "

if /i not "%FTP_URL:~0,8%"=="ftps://" (
    echo ftps:// で始まる URL を指定してください。
    exit /b 1
)

if "%ACTION%"=="1" goto upload
if "%ACTION%"=="2" goto download
echo 操作が選択されていません。
exit /b 1

:upload

set "SOURCE_DIR="
set /p "SOURCE_DIR=アップロードするローカルフォルダ: "

if not exist "%SOURCE_DIR%\." (
    echo フォルダが見つかりません。
    exit /b 1
)

for %%F in ("%SOURCE_DIR%\*") do if exist "%%~fF" if not exist "%%~fF\." curl.exe --ssl-reqd --ftp-pasv --user "%FTP_USER%" --upload-file "%%~fF" "%FTP_URL%"
if errorlevel 1 echo アップロードに失敗しました。

pause
exit /b

:download

set "DEST_DIR="
set /p "DEST_DIR=ダウンロード先フォルダ: "
if not exist "%DEST_DIR%\." mkdir "%DEST_DIR%"
set "FTP_PATTERN="
set /p "FTP_PATTERN=ダウンロード対象 URL（例: ftps://server/path/*.zip）: "

curl.exe --ssl-reqd --ftp-pasv --user "%FTP_USER%" --remote-name --output-dir "%DEST_DIR%" "%FTP_PATTERN%"
if errorlevel 1 echo ダウンロードに失敗しました。

pause
