@echo off
rem 概要: 「自動起動レジストリを点検」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "18_自動起動レジストリを点検.bat"
rem 実行結果（想定結果）: 「自動起動レジストリを点検」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. スクリプト内の設定値を使い、実行条件を確認します。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: Get-ItemPropertyでレジストリ値を取得、Export-CsvでCSVファイルを出力。複数の分岐がある場合は条件に応じて実行します。
rem   3. 結果やログを本文で指定したファイルへ保存し、必要に応じて画面にも表示します。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set "OUTPUT=%~dp0registry_startup_audit.csv"

powershell -NoProfile -Command "$keys=@('HKCU:\Software\Microsoft\Windows\CurrentVersion\Run','HKCU:\Software\Microsoft\Windows\CurrentVersion\RunOnce','HKLM:\Software\Microsoft\Windows\CurrentVersion\Run','HKLM:\Software\Microsoft\Windows\CurrentVersion\RunOnce','HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Run'); $rows=@(); foreach($key in $keys){if(Test-Path $key){$props=Get-ItemProperty -Path $key; foreach($prop in $props.PSObject.Properties | Where-Object {$_.Name -notmatch '^PS'} ){$cmd=[string]$prop.Value; $exe=if($cmd.StartsWith([char]34)){$end=$cmd.IndexOf([char]34,1); if($end -gt 1){$cmd.Substring(1,$end-1)}else{'(unparsed)'}}elseif($cmd -match '^(.+?\.exe)(?:\s|$)'){$matches[1]}else{'(unparsed)'}; $review='review path'; if($exe -match '^%|^C:\\Windows\\|^C:\\Program Files\\'){$review='check publisher'}; $rows += [pscustomobject]@{RegistryKey=$key;Name=$prop.Name;Executable=$exe;Review=$review}}}}; $rows | Export-Csv -LiteralPath $env:OUTPUT -NoTypeInformation -Encoding UTF8; $rows | Format-Table -AutoSize; Write-Host '一覧は参考情報です。正当性は署名元やインストール履歴と照合してください。'"
if errorlevel 1 echo レジストリを点検できませんでした。
echo 出力先: %OUTPUT%

pause
