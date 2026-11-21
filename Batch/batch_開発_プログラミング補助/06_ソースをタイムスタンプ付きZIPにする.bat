@echo off
rem 概要: 「ソースをタイムスタンプ付きZIPにする」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "06_ソースをタイムスタンプ付きZIPにする.bat"
rem 実行結果（想定結果）: 「ソースをタイムスタンプ付きZIPにする」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: Gitでリポジトリを操作。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set "SOURCE="
set "DEST="
set /p "SOURCE=バックアップするソースフォルダ: "
set /p "DEST=ZIP の保存先フォルダ: "

if not exist "%SOURCE%\." (
    echo ソースフォルダが見つかりません。

    pause
    exit /b 1
)

if not exist "%DEST%\." mkdir "%DEST%"

powershell -NoProfile -Command "$s=[IO.Path]::GetFullPath($env:SOURCE).TrimEnd('\'); $d=[IO.Path]::GetFullPath($env:DEST).TrimEnd('\'); if($d -eq $s -or $d.StartsWith($s+'\',[StringComparison]::OrdinalIgnoreCase)){Write-Error 'ZIP 保存先をソースフォルダ内には指定できません'; exit 1}"

if errorlevel 1 (

    pause
    exit /b 1
)

for /f %%T in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd_HHmmss"') do set "STAMP=%%T"
for %%D in ("%SOURCE%") do set "NAME=%%~nD"
set "ARCHIVE=%DEST%\%NAME%_%STAMP%.zip"
tar.exe -a -c -f "%ARCHIVE%" --exclude=.git --exclude=node_modules --exclude=bin --exclude=obj -C "%SOURCE%" .

if errorlevel 1 (
    echo ZIP 作成に失敗しました。

    pause
    exit /b 1
)

echo 作成しました: %ARCHIVE%

pause
