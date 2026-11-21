@echo off
rem 概要: 「ごみ箱を空にする」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "09_ごみ箱を空にする.bat"
rem 実行結果（想定結果）: 「ごみ箱を空にする」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: PowerShellのClear-RecycleBinで対象を処理。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
echo ごみ箱内の項目を完全に削除します。この操作は元に戻せません。
set "CONFIRM="
set /p "CONFIRM=実行する場合は EMPTY と入力してください: "
if not "%CONFIRM%"=="EMPTY" exit /b 0

powershell -NoProfile -Command "$ErrorActionPreference='Stop'; Clear-RecycleBin -Force -ErrorAction Stop"
if errorlevel 1 echo ごみ箱を空にできませんでした。

pause
