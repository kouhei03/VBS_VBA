@echo off
rem 概要: 「メモリ使用量を定期記録」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "09_メモリ使用量を定期記録.bat"
rem 実行結果（想定結果）: 「メモリ使用量を定期記録」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: Set-Contentで結果をファイルへ書き込み、Get-CimInstanceでWindows情報を取得、Add-Contentでログへ追記。複数の分岐がある場合は条件に応じて実行します。
rem   3. 結果やログを本文で指定したファイルへ保存し、必要に応じて画面にも表示します。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set "INTERVAL=60"
set "LOG=%~dp0memory_usage.csv"
set /p "INTERVAL=記録間隔（秒、既定60）: "
if "%INTERVAL%"=="" set "INTERVAL=60"
echo Ctrl+C で停止します。ログ: %LOG%

powershell -NoProfile -Command "$seconds=0; if(-not [int]::TryParse($env:INTERVAL,[ref]$seconds) -or $seconds -lt 1){throw '間隔は1秒以上にしてください'}; if(-not (Test-Path -LiteralPath $env:LOG)){ 'Time,TotalMB,FreeMB,UsedPercent' | Set-Content -LiteralPath $env:LOG -Encoding UTF8 }; while($true){$os=Get-CimInstance Win32_OperatingSystem; $total=[math]::Round($os.TotalVisibleMemorySize/1024,1); $free=[math]::Round($os.FreePhysicalMemory/1024,1); $used=[math]::Round(100*($total-$free)/$total,1); $line='{0},{1},{2},{3}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$total,$free,$used; Add-Content -LiteralPath $env:LOG -Value $line -Encoding UTF8; Write-Host $line; Start-Sleep -Seconds $seconds}"
