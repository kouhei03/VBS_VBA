@echo off
rem 概要: 「ディスク空き容量を監視」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "10_ディスク空き容量を監視.bat"
rem 実行結果（想定結果）: 「ディスク空き容量を監視」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: Get-CimInstanceでWindows情報を取得、Add-Contentでログへ追記。複数の分岐がある場合は条件に応じて実行します。
rem   3. 結果やログを本文で指定したファイルへ保存し、必要に応じて画面にも表示します。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set "DRIVE="
set "LIMIT_GB="
set "LOG=%~dp0disk_warning.log"
set /p "DRIVE=監視するドライブ文字（例: C）: "
set /p "LIMIT_GB=警告する空き容量（GB）: "
echo 60秒ごとに監視します。停止は Ctrl+C です。

powershell -NoProfile -Command "$letter=$env:DRIVE.TrimEnd(':'); $limit=0; if($letter -notmatch '^[A-Za-z]$' -or -not [int]::TryParse($env:LIMIT_GB,[ref]$limit) -or $limit -lt 1){throw 'ドライブ文字または容量を確認してください'}; while($true){$disk=Get-CimInstance Win32_LogicalDisk | Where-Object DeviceID -eq ($letter+':'); if($disk){$free=[math]::Round($disk.FreeSpace/1GB,2); if($free -le $limit){$line='{0} WARNING {1}: 空き {2} GB (閾値 {3} GB)' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$letter,$free,$limit; Add-Content -LiteralPath $env:LOG -Value $line -Encoding UTF8; Write-Warning $line}; Start-Sleep -Seconds 60}else{throw 'ドライブが見つかりません'}}"
