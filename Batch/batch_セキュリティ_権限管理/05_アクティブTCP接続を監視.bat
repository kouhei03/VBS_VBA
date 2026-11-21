@echo off
rem 概要: 「アクティブTCP接続を監視」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "05_アクティブTCP接続を監視.bat"
rem 実行結果（想定結果）: 「アクティブTCP接続を監視」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. スクリプト内の設定値を使い、実行条件を確認します。
rem   2. 主処理では、以下のコマンドを使います: Get-NetTCPConnectionでTCP接続を取得、Get-Processで対象プロセスを取得、Add-Contentでログへ追記。複数の分岐がある場合は条件に応じて実行します。
rem   3. 結果やログを本文で指定したファイルへ保存し、必要に応じて画面にも表示します。

chcp 65001 >nul
setlocal
set "LOG=%~dp0tcp_connections.csv"
echo 時刻,状態,ローカルIP,ローカルポート,リモートIP,リモートポート,PID,プロセス>"%LOG%"
echo 接続一覧を10秒ごとに監視します。停止は Ctrl+C です。

powershell -NoProfile -Command "$seen=@{}; while($true){$items=Get-NetTCPConnection -State Established -ErrorAction SilentlyContinue; foreach($c in $items){$key='{0}|{1}|{2}' -f ($c.LocalAddress+':'+$c.LocalPort),($c.RemoteAddress+':'+$c.RemotePort),$c.OwningProcess; if(-not $seen.ContainsKey($key)){$p=Get-Process -Id $c.OwningProcess -ErrorAction SilentlyContinue; $name=if($p){$p.ProcessName}else{'Unknown'}; $line='{0},{1},{2}:{3},{4}:{5},{6},{7}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$c.State,$c.LocalAddress,$c.LocalPort,$c.RemoteAddress,$c.RemotePort,$c.OwningProcess,$name; Add-Content -LiteralPath $env:LOG -Value $line -Encoding UTF8; Write-Host $line; $seen[$key]=$true}}; Start-Sleep -Seconds 10}"
