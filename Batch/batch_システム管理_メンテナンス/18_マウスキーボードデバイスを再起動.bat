@echo off
rem 概要: 「マウスキーボードデバイスを再起動」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "18_マウスキーボードデバイスを再起動.bat"
rem 実行結果（想定結果）: 「マウスキーボードデバイスを再起動」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. スクリプト内の設定値を使い、実行条件を確認します。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: 管理者権限の有無を確認、PowerShellのGet-PnpDeviceで対象を処理、PowerShellのGet-Commandで対象を処理、PowerShellのRestart-PnpDeviceで対象を処理。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
net session >nul 2>&1

if errorlevel 1 (
    echo 管理者として実行してください。

    pause
    exit /b 1
)

echo 入力デバイスを再起動すると操作できなくなる可能性があります。
echo 対象デバイス名と InstanceId を表示します。

powershell -NoProfile -Command "$devices=Get-PnpDevice -Class Keyboard,Mouse -ErrorAction Stop; $devices | Format-Table Status,Class,FriendlyName,InstanceId -AutoSize; $id=Read-Host '再起動するデバイスの InstanceId を正確に入力'; $d=$devices | Where-Object InstanceId -eq $id; if(-not $d){throw '一致するデバイスがありません'}; $answer=Read-Host '選択したデバイスを再起動する場合は RESTART と入力'; if($answer -cne 'RESTART'){Write-Host '中止しました。'; exit 0}; if(-not (Get-Command Restart-PnpDevice -ErrorAction SilentlyContinue)){throw 'Restart-PnpDevice が利用できません。このスクリプトでは代替の無効化操作を行いません'}; Restart-PnpDevice -InstanceId $id -Confirm:$false -ErrorAction Stop"
if errorlevel 1 echo デバイスを再起動できませんでした。

pause
