@echo off
rem 概要: 「Pingパケットロスを記録」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "20_Pingパケットロスを記録.bat"
rem 実行結果（想定結果）: 「Pingパケットロスを記録」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: 指定ホストへの疎通を確認、Add-Contentでログへ追記。複数の分岐がある場合は条件に応じて実行します。
rem   3. 結果やログを本文で指定したファイルへ保存し、必要に応じて画面にも表示します。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set "HOST="
set /p "HOST=監視するホスト名または IP アドレス: "
if "%HOST%"=="" exit /b 1
set "LOG=%~dp0packet_loss.csv"
echo 時刻,状態>"%LOG%"
echo 停止するには Ctrl+C を押してください。ログ: %LOG%

powershell -NoProfile -Command "while($true){$time=Get-Date -Format 'yyyy-MM-dd HH:mm:ss'; & ping.exe -n 1 -w 1000 $env:HOST *> $null; if($LASTEXITCODE -eq 0){$status='応答'}else{$status='パケットロス'}; Add-Content -LiteralPath $env:LOG -Value ($time+','+$status) -Encoding UTF8; Write-Host ($time+' '+$status); Start-Sleep -Seconds 1}"
