@echo off
rem 概要: 「環境変数の秘密情報名を検査」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "16_環境変数の秘密情報名を検査.bat"
rem 実行結果（想定結果）: 「環境変数の秘密情報名を検査」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. スクリプト内の設定値を使い、実行条件を確認します。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: Export-CsvでCSVファイルを出力。複数の分岐がある場合は条件に応じて実行します。
rem   3. 結果やログを本文で指定したファイルへ保存し、必要に応じて画面にも表示します。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set "OUTPUT=%~dp0environment_secret_findings.csv"
echo 環境変数の値は出力せず、秘密情報らしい変数名だけを記録します。

powershell -NoProfile -Command "$items=@(); foreach($scope in 'Process','User','Machine'){ $vars=[Environment]::GetEnvironmentVariables($scope); foreach($key in $vars.Keys){$name=[string]$key; $value=[string]$vars[$key]; $byName=$name -match '(?i)(API.?KEY|TOKEN|SECRET|PASSWORD|CREDENTIAL|PRIVATE.?KEY)'; $byValue=$value -match '(?i)(AKIA[0-9A-Z]{16}|gh[pousr]_[A-Za-z0-9]{20,}|sk-[A-Za-z0-9_-]{20,})'; if($byName -or $byValue){$reason='secret-like value (redacted)'; if($byName){$reason='sensitive-looking name'}; $items+= [pscustomobject]@{Scope=$scope;Name=$name;Reason=$reason}}}}; $items | Export-Csv -LiteralPath $env:OUTPUT -NoTypeInformation -Encoding UTF8; $items | Format-Table -AutoSize; Write-Host ('検出数: '+$items.Count+'。値は記録していません。')"
if errorlevel 1 echo 検査に失敗しました。
echo 出力先: %OUTPUT%

pause
