@echo off
rem 概要: 「空フォルダを削除」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "10_空フォルダを削除.bat"
rem 実行結果（想定結果）: 「空フォルダを削除」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: Get-ChildItemで対象を列挙、Remove-Itemで対象を削除。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set /p "TARGET_DIR=対象フォルダ（このフォルダ自体は削除しません）: "

if not exist "%TARGET_DIR%\." (
    echo 対象フォルダが見つかりません。

    pause
    exit /b 1
)

powershell -NoProfile -Command "$ErrorActionPreference='Stop'; $dirs=@(Get-ChildItem -LiteralPath $env:TARGET_DIR -Directory -Recurse -Force | Sort-Object { $_.FullName.Length } -Descending); $empty=@(); $planned=@{}; foreach($d in $dirs){$children=@(Get-ChildItem -LiteralPath $d.FullName -Force); $remaining=@($children | Where-Object {-not ($_.PSIsContainer -and $planned.ContainsKey($_.FullName))}); if($remaining.Count -eq 0){$empty+=$d; $planned[$d.FullName]=$true}}; $empty | Select-Object -ExpandProperty FullName; if($empty.Count -eq 0){Write-Host '空フォルダはありません。'} else {$answer=Read-Host '削除する場合は DELETE と入力してください'; if($answer -ceq 'DELETE'){$removed=0; foreach($d in $empty){if(Test-Path -LiteralPath $d.FullName){Remove-Item -LiteralPath $d.FullName -Force; $removed++}}; Write-Host ('削除した空フォルダ数: ' + $removed)}else{Write-Host '中止しました。'}}"
if errorlevel 1 echo フォルダを削除できませんでした。

pause
