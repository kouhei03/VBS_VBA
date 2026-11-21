@echo off
rem 概要: 「エラー時だけ再試行を確認」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "12_エラー時だけ再試行を確認.bat"
rem 実行結果（想定結果）: 「エラー時だけ再試行を確認」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では本文内のWindowsまたはPowerShellコマンドを実行します。分岐がある場合は選択条件に従います。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set "CHILD="
set /p "CHILD=実行する子 BAT ファイルのフルパス: "

if not exist "%CHILD%" (
    echo ファイルが見つかりません。

    pause
    exit /b 1
)

:run_child

call "%CHILD%"
set "RESULT=%ERRORLEVEL%"
if "%RESULT%"=="0" goto success
echo 子 BAT がエラーコード %RESULT% で終了しました。
choice /c RCE /n /m "再試行 (R)、続行 (C)、終了 (E): "
if errorlevel 3 exit /b 1
if errorlevel 2 goto continue_after_error
if errorlevel 1 goto run_child

:continue_after_error

echo エラーを確認して続行しました。
exit /b 0

:success

echo 子 BAT が正常に終了しました。
