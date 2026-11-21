@echo off
rem 概要: 「証明書をインポートまたは公開証明書を出力」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "15_証明書をインポートまたは公開証明書を出力.bat"
rem 実行結果（想定結果）: 「証明書をインポートまたは公開証明書を出力」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: Import-Certificateで証明書を取り込み、Get-ChildItemで対象を列挙、Export-Certificateで証明書を出力。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
echo 1: .cer 公開証明書を CurrentUser\My にインポート
echo 2: ストアから公開証明書のみを .cer に出力（秘密鍵は出力しません）
set "ACTION="
set /p "ACTION=操作を選択してください (1/2): "
if "%ACTION%"=="1" goto import_cert
if "%ACTION%"=="2" goto export_cert
exit /b 1

:import_cert

set "CERT="
set /p "CERT=信頼できる発行元から取得した .cer ファイル: "

if not exist "%CERT%" (
    echo 証明書ファイルが見つかりません。

    pause
    exit /b 1
)

powershell -NoProfile -Command "Import-Certificate -FilePath $env:CERT -CertStoreLocation Cert:\CurrentUser\My -ErrorAction Stop | Format-List Subject,Thumbprint,NotAfter"
if errorlevel 1 echo インポートに失敗しました。

pause
exit /b

:export_cert

set "THUMBPRINT="
set "OUTPUT="
set /p "THUMBPRINT=CurrentUser\My の証明書 Thumbprint: "
set /p "OUTPUT=出力する .cer ファイルパス: "

powershell -NoProfile -Command "$cert=Get-ChildItem Cert:\CurrentUser\My | Where-Object Thumbprint -eq ($env:THUMBPRINT -replace '\s','') | Select-Object -First 1; if(-not $cert){throw '証明書が見つかりません'}; if(Test-Path -LiteralPath $env:OUTPUT){throw '出力先が既に存在します'}; Export-Certificate -Cert $cert -FilePath $env:OUTPUT -Type CERT -ErrorAction Stop | Out-Null; Write-Host '公開証明書のみ出力しました'"
if errorlevel 1 echo 証明書を出力できませんでした。

pause
