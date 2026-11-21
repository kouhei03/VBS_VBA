@echo off
rem 概要: 「FreeFileSyncで双方向同期」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "13_FreeFileSyncで双方向同期.bat"
rem 実行結果（想定結果）: 「FreeFileSyncで双方向同期」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では本文内のWindowsまたはPowerShellコマンドを実行します。分岐がある場合は選択条件に従います。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。

chcp 65001 >nul
setlocal
set /p "FFS=FreeFileSync.exe のフルパス: "
set /p "CONFIG=事前に作成した双方向同期 .ffs_gui 設定ファイル: "

if not exist "%FFS%" (
    echo FreeFileSync.exe が見つかりません。

    pause
    exit /b 1
)

if not exist "%CONFIG%" (
    echo 同期設定ファイルが見つかりません。

    pause
    exit /b 1
)

echo 同期設定の内容を確認してから実行してください。削除伝播が設定されている場合はご注意ください。
set /p "CONFIRM=FreeFileSync を起動しますか？ (Y/N): "
if /i not "%CONFIRM%"=="Y" exit /b 0
start "" "%FFS%" "%CONFIG%"
