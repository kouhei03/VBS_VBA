@echo off
rem 概要: 「子BATの実行時間を計測」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "18_子BATの実行時間を計測.bat"
rem 実行結果（想定結果）: 「子BATの実行時間を計測」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では本文内のWindowsまたはPowerShellコマンドを実行します。分岐がある場合は選択条件に従います。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。

chcp 65001 >nul
setlocal
set "CHILD="
set /p "CHILD=計測する子 BAT ファイルのフルパス: "

if not exist "%CHILD%" (
    echo ファイルが見つかりません。

    pause
    exit /b 1
)

for /f %%T in ('powershell -NoProfile -Command "[DateTimeOffset]::UtcNow.ToUnixTimeSeconds()"') do set "STARTTIME=%%T"
call "%CHILD%"
set "RESULT=%ERRORLEVEL%"
for /f %%T in ('powershell -NoProfile -Command "[DateTimeOffset]::UtcNow.ToUnixTimeSeconds()"') do set "ENDTIME=%%T"
set /a ELAPSED=ENDTIME-STARTTIME
set /a MINUTES=ELAPSED/60
set /a SECONDS=ELAPSED%%60
echo 実行時間: %MINUTES% 分 %SECONDS% 秒
exit /b %RESULT%
