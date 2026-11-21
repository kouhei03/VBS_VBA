@echo off
rem 概要: 「AWSまたはAzureのVMを起動停止」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "01_AWSまたはAzureのVMを起動停止.bat"
rem 実行結果（想定結果）: 「AWSまたはAzureのVMを起動停止」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: 指定したAWS EC2インスタンスを起動、指定したAWS EC2インスタンスを停止、指定したAzure VMを起動、指定したAzure VMの割り当てを解除。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
echo 操作対象を AWS EC2 または Azure VM の1台に限定します。
echo 1: AWS EC2 起動
echo 2: AWS EC2 停止
echo 3: Azure VM 起動
echo 4: Azure VM 停止（割り当て解除）
set "ACTION="
set /p "ACTION=操作を選択してください (1-4): "
set "REGION="
set "INSTANCE="
set "RESOURCE_GROUP="
set "VM="
if "%ACTION%"=="1" goto aws_start
if "%ACTION%"=="2" goto aws_stop
if "%ACTION%"=="3" goto azure_start
if "%ACTION%"=="4" goto azure_stop
exit /b 1

:aws_start

set /p "INSTANCE=対象 EC2 インスタンス ID: "
set /p "REGION=AWS リージョン（例: ap-northeast-1）: "
set "CONFIRM="
set /p "CONFIRM=指定した EC2 を起動する場合は START と入力: "
if not "%CONFIRM%"=="START" exit /b 0

aws ec2 start-instances --instance-ids "%INSTANCE%" --region "%REGION%"
goto result

:aws_stop

set /p "INSTANCE=対象 EC2 インスタンス ID: "
set /p "REGION=AWS リージョン（例: ap-northeast-1）: "
set "CONFIRM="
set /p "CONFIRM=指定した EC2 を停止する場合は STOP と入力: "
if not "%CONFIRM%"=="STOP" exit /b 0

aws ec2 stop-instances --instance-ids "%INSTANCE%" --region "%REGION%"
goto result

:azure_start

set /p "RESOURCE_GROUP=Azure リソースグループ名: "
set /p "VM=対象 VM 名: "
set "CONFIRM="
set /p "CONFIRM=指定した Azure VM を起動する場合は START と入力: "
if not "%CONFIRM%"=="START" exit /b 0

az vm start --resource-group "%RESOURCE_GROUP%" --name "%VM%"
goto result

:azure_stop

set /p "RESOURCE_GROUP=Azure リソースグループ名: "
set /p "VM=対象 VM 名: "
set "CONFIRM="
set /p "CONFIRM=指定した Azure VM を割り当て解除する場合は DEALLOCATE と入力: "
if not "%CONFIRM%"=="DEALLOCATE" exit /b 0

az vm deallocate --resource-group "%RESOURCE_GROUP%" --name "%VM%"

:result

if errorlevel 1 echo 操作に失敗しました。CLI の認証、権限、対象IDを確認してください。

pause
