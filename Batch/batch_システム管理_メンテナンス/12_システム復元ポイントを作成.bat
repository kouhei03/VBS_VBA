@echo off
rem 概要: 「システム復元ポイントを作成」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "12_システム復元ポイントを作成.bat"
rem 実行結果（想定結果）: 「システム復元ポイントを作成」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. スクリプト内の設定値を使い、実行条件を確認します。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: 管理者権限の有無を確認、Checkpoint-Computerでシステム復元ポイントを作成。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
net session >nul 2>&1

if errorlevel 1 (
    echo 管理者として実行してください。

    pause
    exit /b 1
)

echo 復元ポイント作成にはシステムの保護が有効である必要があります。

powershell -NoProfile -Command "$ErrorActionPreference='Stop'; Checkpoint-Computer -Description ('手動復元ポイント '+(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')) -RestorePointType MODIFY_SETTINGS"
if errorlevel 1 echo 復元ポイントを作成できませんでした。システムの保護設定を確認してください。

pause
