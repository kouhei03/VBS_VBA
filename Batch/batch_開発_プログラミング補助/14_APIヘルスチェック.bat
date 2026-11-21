@echo off
rem 概要: 「APIヘルスチェック」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "14_APIヘルスチェック.bat"
rem 実行結果（想定結果）: 「APIヘルスチェック」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: Invoke-WebRequestでHTTP応答を取得、Export-CsvでCSVファイルを出力。複数の分岐がある場合は条件に応じて実行します。
rem   3. 結果やログを本文で指定したファイルへ保存し、必要に応じて画面にも表示します。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set "URL1="
set "URL2="
set "URL3="
set "URL4="
set "URL5="
set "OUTPUT=%~dp0health_check.csv"
set /p "URL1=1つ目の HTTP(S) エンドポイント: "
set /p "URL2=2つ目（省略可）: "
set /p "URL3=3つ目（省略可）: "
set /p "URL4=4つ目（省略可）: "
set /p "URL5=5つ目（省略可）: "

powershell -NoProfile -Command "$urls=@($env:URL1,$env:URL2,$env:URL3,$env:URL4,$env:URL5) | Where-Object {$_}; $results=@(foreach($url in $urls){$sw=[Diagnostics.Stopwatch]::StartNew(); try{if($url -notmatch '^https?://'){throw 'URL は http:// または https:// で指定してください'}; $r=Invoke-WebRequest -Uri $url -TimeoutSec 10 -UseBasicParsing; $code=[int]$r.StatusCode; $state=if($code -ge 200 -and $code -lt 400){'OK'}else{'NG'}}catch{$code=0; $state='NG'}; $sw.Stop(); [pscustomobject]@{Time=Get-Date -Format 'yyyy-MM-dd HH:mm:ss'; Url=$url; Status=$code; Result=$state; Milliseconds=$sw.ElapsedMilliseconds}}); $results | Format-Table -AutoSize; if($results.Count){if(Test-Path -LiteralPath $env:OUTPUT){$results | Export-Csv -LiteralPath $env:OUTPUT -NoTypeInformation -Append -Encoding UTF8}else{$results | Export-Csv -LiteralPath $env:OUTPUT -NoTypeInformation -Encoding UTF8}}"

pause
