@echo off
rem 概要: 「多重起動を防ぐロック」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "17_多重起動を防ぐロック.bat"
rem 実行結果（想定結果）: 「多重起動を防ぐロック」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. スクリプト内の設定値を使い、実行条件を確認します。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: 指定秒数の待機を実行。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set "LOCKDIR=%~dp0.task_lock_17"
mkdir "%LOCKDIR%" 2>nul

if errorlevel 1 (
    echo 既に実行中か、ロックが残っています: %LOCKDIR%
    echo 異常終了後は実行中でないことを確認してから、このロックフォルダを手動で削除してください。
    exit /b 1
)

echo ロックを取得しました。作業中は Ctrl+C で中断しないでください。
call :work
set "RESULT=%ERRORLEVEL%"
rmdir "%LOCKDIR%"

if errorlevel 1 (
    echo ロックを解除できませんでした: %LOCKDIR%
    exit /b 1
)

echo ロックを解除しました。
exit /b %RESULT%

:work

echo 多重起動を防ぎたい処理をこのサブルーチン内に記述してください。

timeout /t 5 /nobreak >nul
exit /b 0
