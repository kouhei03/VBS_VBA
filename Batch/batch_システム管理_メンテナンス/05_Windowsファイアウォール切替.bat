@echo off
rem 概要: 「Windowsファイアウォール切替」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "05_Windowsファイアウォール切替.bat"
rem 実行結果（想定結果）: 「Windowsファイアウォール切替」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: 管理者権限の有無を確認、Windowsネットワーク設定を参照または変更。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
net session >nul 2>&1

if errorlevel 1 (
    echo 管理者として実行してください。

    pause
    exit /b 1
)

echo 1: 有効化
echo 2: 一時的に無効化（保護が弱くなります）
set "ACTION="
set /p "ACTION=操作を選択してください (1/2): "
if "%ACTION%"=="1" goto enable_firewall
if "%ACTION%"=="2" goto disable_firewall
echo 操作が選択されていません。
exit /b 1

:enable_firewall

netsh advfirewall set allprofiles state on
goto done

:disable_firewall

set "CONFIRM="
set /p "CONFIRM=全プロファイルを無効化します。実行する場合は OFF と入力してください: "
if /i not "%CONFIRM%"=="OFF" exit /b 0

netsh advfirewall set allprofiles state off

:done

pause
