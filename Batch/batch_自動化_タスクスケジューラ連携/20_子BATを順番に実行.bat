@echo off
rem 概要: 「子BATを順番に実行」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "20_子BATを順番に実行.bat"
rem 実行結果（想定結果）: 「子BATを順番に実行」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では本文内のWindowsまたはPowerShellコマンドを実行します。分岐がある場合は選択条件に従います。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set "CHILD1="
set "CHILD2="
set "CHILD3="
set /p "CHILD1=1番目に実行する子 BAT のフルパス: "
set /p "CHILD2=2番目に実行する子 BAT のフルパス: "
set /p "CHILD3=3番目に実行する子 BAT のフルパス: "

if not exist "%CHILD1%" (
    echo 1番目のファイルが見つかりません。
    exit /b 1
)

if not exist "%CHILD2%" (
    echo 2番目のファイルが見つかりません。
    exit /b 1
)

if not exist "%CHILD3%" (
    echo 3番目のファイルが見つかりません。
    exit /b 1
)

call "%CHILD1%"
if errorlevel 1 goto child_failed
call "%CHILD2%"
if errorlevel 1 goto child_failed
call "%CHILD3%"
if errorlevel 1 goto child_failed
echo すべての子 BAT が正常に終了しました。
exit /b 0

:child_failed

echo 子 BAT のいずれかがエラー終了したため、後続処理を中止しました。
exit /b 1
