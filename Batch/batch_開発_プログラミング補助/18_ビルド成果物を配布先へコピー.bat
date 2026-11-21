@echo off
rem 概要: 「ビルド成果物を配布先へコピー」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "18_ビルド成果物を配布先へコピー.bat"
rem 実行結果（想定結果）: 「ビルド成果物を配布先へコピー」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: ファイルをコピー、Robocopyでファイルをコピーまたは同期。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set "SOURCE="
set "DEST="
set /p "SOURCE=ビルド成果物フォルダ: "
set /p "DEST=配布先フォルダ: "

if not exist "%SOURCE%\." (
    echo ビルド成果物フォルダが見つかりません。

    pause
    exit /b 1
)

if not exist "%DEST%\." mkdir "%DEST%"

powershell -NoProfile -Command "$s=[IO.Path]::GetFullPath($env:SOURCE).TrimEnd('\'); $d=[IO.Path]::GetFullPath($env:DEST).TrimEnd('\'); if($d -eq $s -or $d.StartsWith($s+'\',[StringComparison]::OrdinalIgnoreCase)){Write-Error '配布先を成果物フォルダ内には指定できません'; exit 1}"

if errorlevel 1 (

    pause
    exit /b 1
)

echo 既存の配布先ファイルは削除しませんが、同名ファイルは更新されます。
set "CONFIRM="
set /p "CONFIRM=コピーする場合は COPY と入力してください: "
if not "%CONFIRM%"=="COPY" exit /b 0

robocopy "%SOURCE%" "%DEST%" /E /R:2 /W:2
if errorlevel 8 echo コピーに失敗しました。終了コード: %ERRORLEVEL%

pause
