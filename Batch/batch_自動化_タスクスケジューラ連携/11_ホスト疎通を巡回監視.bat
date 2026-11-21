@echo off
rem 概要: 「ホスト疎通を巡回監視」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "11_ホスト疎通を巡回監視.bat"
rem 実行結果（想定結果）: 「ホスト疎通を巡回監視」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: Test-Connectionでホスト疎通を確認、Add-Contentでログへ追記、指定秒数の待機を実行。複数の分岐がある場合は条件に応じて実行します。
rem   3. 結果やログを本文で指定したファイルへ保存し、必要に応じて画面にも表示します。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set "HOST="
set "LOG=%~dp0host_monitor.log"
set /p "HOST=監視するホスト名またはIPアドレス: "
if "%HOST%"=="" exit /b 1
echo 監視を開始します。記録先: %LOG% （停止は Ctrl+C）

powershell -NoProfile -Command "if($env:HOST -notmatch '^[A-Za-z0-9.:-]+$'){throw 'ホスト名またはIPアドレス以外の文字が含まれています'}"
if errorlevel 1 exit /b 1

:monitor

powershell -NoProfile -Command "$ok=Test-Connection -ComputerName $env:HOST -Count 1 -Quiet -ErrorAction SilentlyContinue; $status='no reply'; if($ok){$status='reachable'}; $line='{0} - {1} {2}' -f (Get-Date -Format s),$status,$env:HOST; Add-Content -LiteralPath $env:LOG -Value $line -Encoding UTF8; Write-Host $line"

timeout /t 30 /nobreak >nul
goto monitor
