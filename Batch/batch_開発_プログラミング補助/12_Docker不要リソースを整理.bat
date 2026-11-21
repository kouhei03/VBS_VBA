@echo off
rem 概要: 「Docker不要リソースを整理」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "12_Docker不要リソースを整理.bat"
rem 実行結果（想定結果）: 「Docker不要リソースを整理」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: Dockerの未使用リソースを削除。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。

chcp 65001 >nul
setlocal
echo 削除対象の容量を確認します:

docker system df
echo.
echo docker system prune は未使用コンテナ、ネットワーク、dangling イメージ、ビルドキャッシュを削除します。
echo ボリュームと全ての未使用イメージは対象にしません。
set "CONFIRM="
set /p "CONFIRM=続行する場合は PRUNE と入力してください: "
if not "%CONFIRM%"=="PRUNE" exit /b 0

docker system prune

pause
