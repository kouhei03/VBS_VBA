@echo off
rem 概要: 「APIを定期ヘルスチェック」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "12_APIを定期ヘルスチェック.bat"
rem 実行結果（想定結果）: 「APIを定期ヘルスチェック」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: Invoke-WebRequestでHTTP応答を取得、Add-Contentでログへ追記。複数の分岐がある場合は条件に応じて実行します。
rem   3. 結果やログを本文で指定したファイルへ保存し、必要に応じて画面にも表示します。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set "URL="
set "INTERVAL="
set "COUNT="
set "LOG=%~dp0api_health.log"
set /p "URL=確認する HTTPS URL: "
set /p "INTERVAL=確認間隔（秒、30-3600）: "
set /p "COUNT=実行回数（1-50）: "

powershell -NoProfile -Command "if($env:URL -notmatch '^https://[^ ]+$' -or $env:INTERVAL -notmatch '^\d+$' -or [int]$env:INTERVAL -lt 30 -or [int]$env:INTERVAL -gt 3600 -or $env:COUNT -notmatch '^\d+$' -or [int]$env:COUNT -lt 1 -or [int]$env:COUNT -gt 50){throw 'HTTPS URL、間隔、回数を確認してください'}"
if errorlevel 1 exit /b 1
echo 指定URLへ GET を最大 %COUNT% 回送信します。ログ: %LOG%

powershell -NoProfile -Command "for($i=1;$i -le [int]$env:COUNT;$i++){try{$response=Invoke-WebRequest -Uri $env:URL -Method Get -TimeoutSec 20 -UseBasicParsing; $code=[int]$response.StatusCode}catch{$code=if($_.Exception.Response){[int]$_.Exception.Response.StatusCode}else{0}}; $line='{0} attempt={1} status={2}' -f (Get-Date -Format s),$i,$code; Add-Content -LiteralPath $env:LOG -Value $line -Encoding UTF8; Write-Host $line; if($i -lt [int]$env:COUNT){Start-Sleep -Seconds ([int]$env:INTERVAL)}}"
if errorlevel 1 echo ヘルスチェックを完了できませんでした。

pause
