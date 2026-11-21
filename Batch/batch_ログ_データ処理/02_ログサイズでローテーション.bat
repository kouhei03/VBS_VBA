@echo off
rem 概要: 「ログサイズでローテーション」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "02_ログサイズでローテーション.bat"
rem 実行結果（想定結果）: 「ログサイズでローテーション」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: PowerShellのGet-Itemで対象を処理、Move-Itemでファイルを移動、ファイルを移動。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set "LOG="
set "LIMIT_MB="
set /p "LOG=監視するログファイル: "
set /p "LIMIT_MB=ローテーションするサイズ（MB）: "
if "%LOG%"=="" exit /b 1
echo 監視中です。上限を超えるとタイムスタンプ付きの名前に変更します。停止は Ctrl+C です。

powershell -NoProfile -Command "$limit=0; if(-not [int]::TryParse($env:LIMIT_MB,[ref]$limit) -or $limit -lt 1){throw 'サイズは1以上の整数にしてください'}; $bytes=[int64]$limit*1MB; while($true){if(Test-Path -LiteralPath $env:LOG){$f=Get-Item -LiteralPath $env:LOG; if($f.Length -ge $bytes){$stamp=Get-Date -Format 'yyyyMMdd_HHmmss'; $dest=$env:LOG+'.'+$stamp; Move-Item -LiteralPath $env:LOG -Destination $dest -ErrorAction Stop; Write-Host ('ローテーション: '+$dest)}}; Start-Sleep -Seconds 5}"
