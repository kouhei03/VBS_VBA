@echo off
rem 概要: 「文字列をSHA256ハッシュ化」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "15_文字列をSHA256ハッシュ化.bat"
rem 実行結果（想定結果）: 「文字列をSHA256ハッシュ化」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。
rem   2. 主処理では本文内のWindowsまたはPowerShellコマンドを実行します。分岐がある場合は選択条件に従います。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。

chcp 65001 >nul
setlocal
echo 入力文字列は画面に表示されます。パスワード等の秘密情報を入力しないでください。
set "TEXT="
set /p "TEXT=ハッシュ化する文字列: "

powershell -NoProfile -Command "$bytes=[Text.Encoding]::UTF8.GetBytes($env:TEXT); $sha=[Security.Cryptography.SHA256]::Create(); try{$hash=[BitConverter]::ToString($sha.ComputeHash($bytes)).Replace('-','').ToLowerInvariant()}finally{$sha.Dispose()}; Write-Host ('SHA-256: '+$hash)"

pause
