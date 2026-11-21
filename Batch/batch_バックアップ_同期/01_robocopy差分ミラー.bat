@echo off
rem 概要: 「robocopy差分ミラー」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "01_robocopy差分ミラー.bat"
rem 実行結果（想定結果）: 「robocopy差分ミラー」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: Robocopyの /MIR でミラー先を同期し、先にだけあるファイルを削除。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set /p "SOURCE=バックアップ元フォルダ: "
set /p "DEST=ミラー先フォルダ（余分なファイルは削除されます）: "

if not exist "%SOURCE%\." (
    echo バックアップ元が見つかりません。

    pause
    exit /b 1
)

echo /MIR はミラー先にだけ存在するファイルを削除します。
set /p "CONFIRM=内容を理解した上で実行する場合は MIRROR と入力してください: "
if not "%CONFIRM%"=="MIRROR" exit /b 0

powershell -NoProfile -Command "$s=[IO.Path]::GetFullPath($env:SOURCE).TrimEnd('\'); $d=[IO.Path]::GetFullPath($env:DEST).TrimEnd('\'); if($s -eq $d -or $s.StartsWith($d+'\',[StringComparison]::OrdinalIgnoreCase) -or $d.StartsWith($s+'\',[StringComparison]::OrdinalIgnoreCase)){Write-Error '元と先は別の階層に指定してください'; exit 1}"

if errorlevel 1 (

    pause
    exit /b 1
)

if not exist "%DEST%\." mkdir "%DEST%"

robocopy "%SOURCE%" "%DEST%" /MIR /R:2 /W:2 /XJ

if errorlevel 8 (
    echo ミラーリングに失敗しました。Robocopy の終了コード: %ERRORLEVEL%
) else (
    echo ミラーリングが完了しました。
)

pause
