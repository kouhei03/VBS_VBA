@echo off
rem 概要: 「BitLocker状態確認と解除」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "11_BitLocker状態確認と解除.bat"
rem 実行結果（想定結果）: 「BitLocker状態確認と解除」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: Unlock-BitLockerでBitLockerボリュームを解除。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set "DRIVE="

manage-bde -status
echo 解除する場合は BitLocker 回復パスワードではなく、通常のドライブ パスワードを使用します。
set "ACTION="
set /p "ACTION=解除する場合は 1、状態確認のみなら Enter: "
if not "%ACTION%"=="1" exit /b 0
set /p "DRIVE=解除するドライブ（例: E:）: "

powershell -NoProfile -Command "$d=$env:DRIVE.TrimEnd('\'); if($d -notmatch '^[A-Za-z]:$'){throw 'ドライブ文字を指定してください'}; $secure=Read-Host 'BitLocker パスワードを入力' -AsSecureString; Unlock-BitLocker -MountPoint $d -Password $secure -ErrorAction Stop"
if errorlevel 1 echo 解除できませんでした。パスワード、ドライブ、BitLocker 状態を確認してください。

pause
