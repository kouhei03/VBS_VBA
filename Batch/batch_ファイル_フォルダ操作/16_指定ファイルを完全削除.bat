@echo off
rem 概要: 「指定ファイルを完全削除」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "16_指定ファイルを完全削除.bat"
rem 実行結果（想定結果）: 「指定ファイルを完全削除」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: PowerShellのGet-Itemで対象を処理、Remove-Itemで対象を削除。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set /p "FILE_PATH=完全削除するファイルのフルパス: "

if not exist "%FILE_PATH%" (
    echo ファイルが見つかりません。

    pause
    exit /b 1
)

echo 次のファイルをごみ箱を経由せず削除します:
echo %FILE_PATH%
echo この操作は元に戻せません。
set /p "CONFIRM=実行する場合は DELETE と入力してください: "
if not "%CONFIRM%"=="DELETE" exit /b 0

powershell -NoProfile -Command "$ErrorActionPreference='Stop'; if((Get-Item -LiteralPath $env:FILE_PATH).PSIsContainer){throw 'フォルダは対象外です'}; Remove-Item -LiteralPath $env:FILE_PATH -Force"
if errorlevel 1 echo ファイルを削除できませんでした。

pause
