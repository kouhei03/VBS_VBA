@echo off
rem 概要: 「成否をログに記録」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "08_成否をログに記録.bat"
rem 実行結果（想定結果）: 「成否をログに記録」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: Robocopyでファイルをコピーまたは同期。複数の分岐がある場合は条件に応じて実行します。
rem   3. 結果やログを本文で指定したファイルへ保存し、必要に応じて画面にも表示します。

chcp 65001 >nul
setlocal
set /p "SOURCE=バックアップ元フォルダ: "
set /p "DEST=バックアップ先フォルダ: "
set /p "LOG=ログファイルのフルパス: "
for %%P in ("%LOG%") do if not exist "%%~dpP." mkdir "%%~dpP"
echo [%date% %time%] Backup started >> "%LOG%"

robocopy "%SOURCE%" "%DEST%" /E /Z /R:3 /W:5 /XJ /LOG+:"%LOG%"
set "RC=%ERRORLEVEL%"

if %RC% GEQ 8 (
    echo [%date% %time%] FAILED: Robocopy exit code %RC% >> "%LOG%"
    echo バックアップ失敗。ログ: %LOG%
    exit /b %RC%
)

echo [%date% %time%] SUCCESS: Robocopy exit code %RC% >> "%LOG%"
echo バックアップ完了。ログ: %LOG%

pause
