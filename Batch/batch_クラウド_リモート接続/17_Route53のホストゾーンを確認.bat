@echo off
rem 概要: 「Route53のホストゾーンを確認」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "17_Route53のホストゾーンを確認.bat"
rem 実行結果（想定結果）: 「Route53のホストゾーンを確認」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: Route 53のホストゾーンを一覧。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
echo AWS アカウントで参照可能な Route 53 ホストゾーン一覧を表示します。
set "REGION="
set /p "REGION=AWS CLI リージョン（任意、空欄可）: "

if "%REGION%"=="" (
    aws route53 list-hosted-zones --query "HostedZones[].{Name:Name,Id:Id,Private:Config.PrivateZone}" --output table
) else (
    aws route53 list-hosted-zones --region "%REGION%" --query "HostedZones[].{Name:Name,Id:Id,Private:Config.PrivateZone}" --output table
)

if errorlevel 1 echo Route 53 の情報を取得できませんでした。

pause
