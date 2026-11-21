@echo off
rem 概要: 「入力内容で処理を分岐」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "10_入力内容で処理を分岐.bat"
rem 実行結果（想定結果）: 「入力内容で処理を分岐」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では本文内のWindowsまたはPowerShellコマンドを実行します。分岐がある場合は選択条件に従います。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
echo 1: サンプル処理 A
echo 2: サンプル処理 B
echo 3: 終了
choice /c 123 /n /m "番号を選択してください (1/2/3): "
if errorlevel 3 goto finish
if errorlevel 2 goto option_b
if errorlevel 1 goto option_a

:option_a

set "VALUE="
set /p "VALUE=処理 A に渡す値: "
echo 処理 A: %VALUE%
goto finish

:option_b

set "VALUE="
set /p "VALUE=処理 B に渡す値: "
echo 処理 B: %VALUE%

:finish

echo 終了します。
