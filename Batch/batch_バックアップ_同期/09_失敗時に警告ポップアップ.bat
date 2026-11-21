@echo off
rem 概要: 「失敗時に警告ポップアップ」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "09_失敗時に警告ポップアップ.bat"
rem 実行結果（想定結果）: 「失敗時に警告ポップアップ」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: Robocopyでファイルをコピーまたは同期、PowerShellのNew-Objectで対象を処理。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。

chcp 65001 >nul
setlocal
set /p "SOURCE=バックアップ元フォルダ: "
set /p "DEST=バックアップ先フォルダ: "

if not exist "%SOURCE%\." (
    echo バックアップ元が見つかりません。

    pause
    exit /b 1
)

robocopy "%SOURCE%" "%DEST%" /E /Z /R:3 /W:5 /XJ
set "RC=%ERRORLEVEL%"

if %RC% GEQ 8 (
    powershell -NoProfile -Command "$w=New-Object -ComObject WScript.Shell; [void]$w.Popup('バックアップに失敗しました。Robocopy 終了コード: %RC%',0,'バックアップ警告',48)"
    exit /b %RC%
)

echo バックアップ完了（Robocopy 終了コード: %RC%）。

pause
