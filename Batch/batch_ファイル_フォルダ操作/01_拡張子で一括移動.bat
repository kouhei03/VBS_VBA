@echo off
rem 概要: 「拡張子で一括移動」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "01_拡張子で一括移動.bat"
rem 実行結果（想定結果）: 「拡張子で一括移動」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: ファイルを移動。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set /p "SOURCE=移動元フォルダ: "
set /p "DEST=移動先フォルダ: "
set /p "EXT=対象の拡張子（例: txt、ドットなし）: "

if not exist "%SOURCE%\." (
    echo 移動元フォルダが見つかりません。

    pause
    exit /b 1
)

if not exist "%DEST%\." mkdir "%DEST%"
echo %SOURCE% 内の *.%EXT% を %DEST% へ移動します。
set /p "CONFIRM=続行しますか？ (Y/N): "
if /i not "%CONFIRM%"=="Y" exit /b 0
move /-Y "%SOURCE%\*.%EXT%" "%DEST%\"
if errorlevel 1 echo 移動できないファイルがあるか、該当ファイルがありません。

pause
