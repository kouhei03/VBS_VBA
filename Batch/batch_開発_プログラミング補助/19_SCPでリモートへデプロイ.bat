@echo off
rem 概要: 「SCPでリモートへデプロイ」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "19_SCPでリモートへデプロイ.bat"
rem 実行結果（想定結果）: 「SCPでリモートへデプロイ」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: SCPでファイルをリモート転送、OpenSSHでリモート接続またはコマンド実行。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
where scp.exe >nul 2>&1

if errorlevel 1 (
    echo OpenSSH Client（scp.exe）が見つかりません。

    pause
    exit /b 1
)

set "SOURCE="
set "TARGET="
set /p "SOURCE=転送するビルド成果物フォルダ: "
set /p "TARGET=SSH 転送先（例: user@server:/var/www/app）: "

if not exist "%SOURCE%\." (
    echo 転送元フォルダが見つかりません。

    pause
    exit /b 1
)

echo SSH 認証を使用します。パスワードや秘密鍵を BAT に書かないでください。
set "CONFIRM="
set /p "CONFIRM=指定先に転送する場合は DEPLOY と入力してください: "
if not "%CONFIRM%"=="DEPLOY" exit /b 0

scp.exe -r "%SOURCE%" "%TARGET%"
if errorlevel 1 echo 転送に失敗しました。SSH 接続、権限、転送先を確認してください。

pause
