@echo off
rem 概要: 「EC2を確認付きで再構築」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "20_EC2を確認付きで再構築.bat"
rem 実行結果（想定結果）: 「EC2を確認付きで再構築」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: 指定したAWS EC2インスタンスを終了、EC2の終了完了を待機、起動テンプレートからEC2を作成。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
echo 警告: 指定した EC2 インスタンスを終了すると、インスタンスストア上のデータは失われます。
echo 既存の起動テンプレートから新しいインスタンスを1台作成します。
set "INSTANCE="
set "TEMPLATE="
set "REGION="
set /p "INSTANCE=終了対象の EC2 インスタンス ID: "
set /p "TEMPLATE=再作成に使う起動テンプレート ID（lt-...）: "
set /p "REGION=AWS リージョン: "

powershell -NoProfile -Command "if($env:INSTANCE -notmatch '^i-[0-9a-f]{8,17}$' -or $env:TEMPLATE -notmatch '^lt-[0-9a-f]{8,17}$' -or $env:REGION -notmatch '^[a-z]{2}(-gov)?-[a-z]+-\d$'){throw 'インスタンスID、テンプレートID、またはリージョンが不正です'}"
if errorlevel 1 exit /b 1

powershell -NoProfile -Command "$result=& aws ec2 terminate-instances --instance-ids $env:INSTANCE --region $env:REGION --dry-run 2>&1; if(($result -join ' ') -notmatch 'DryRunOperation'){Write-Error ($result -join ' '); exit 1}; Write-Host 'AWS dry-run が成功しました。'; $confirm=Read-Host ('終了して起動テンプレート '+$env:TEMPLATE+' から再作成するには、対象ID '+$env:INSTANCE+' をそのまま入力'); if($confirm -cne $env:INSTANCE){Write-Host 'キャンセルしました'; exit 0}; & aws ec2 terminate-instances --instance-ids $env:INSTANCE --region $env:REGION; if($LASTEXITCODE -ne 0){exit $LASTEXITCODE}; & aws ec2 wait instance-terminated --instance-ids $env:INSTANCE --region $env:REGION; if($LASTEXITCODE -ne 0){Write-Error '終了待機に失敗しました'; exit $LASTEXITCODE}; & aws ec2 run-instances --launch-template LaunchTemplateId=$env:TEMPLATE --min-count 1 --max-count 1 --region $env:REGION; if($LASTEXITCODE -ne 0){Write-Error '新しいインスタンスの作成に失敗しました'; exit $LASTEXITCODE}"
if errorlevel 1 echo 再構築処理に失敗しました。AWS コンソールで対象リソースの状態を確認してください。

pause
