@echo off
rem 概要: 「一時ファイルを検索して削除」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "09_一時ファイルを検索して削除.bat"
rem 実行結果（想定結果）: 「一時ファイルを検索して削除」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: Get-ChildItemで対象を列挙、Remove-Itemで対象を削除。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
echo 安全のため、全ディスクではなく現在のユーザーと Windows の一時フォルダだけを対象にします。
echo %TEMP%
echo %WINDIR%\Temp
set /p "CONFIRM=上記の場所にある .tmp ファイルを検索して削除します。続行しますか？ (Y/N): "
if /i not "%CONFIRM%"=="Y" exit /b 0

powershell -NoProfile -Command "$ErrorActionPreference='Stop'; $roots=@($env:TEMP,(Join-Path $env:WINDIR 'Temp')) | Where-Object {Test-Path -LiteralPath $_}; $items=@(foreach($r in $roots){Get-ChildItem -LiteralPath $r -File -Recurse -Filter '*.tmp' -Force}); $items | Select-Object -ExpandProperty FullName; if($items.Count -eq 0){Write-Host '対象ファイルはありません。'} else {$answer=Read-Host '削除する場合は DELETE と入力してください'; if($answer -ceq 'DELETE'){foreach($f in $items){Remove-Item -LiteralPath $f.FullName -Force}; Write-Host ('削除数: ' + $items.Count)}else{Write-Host '中止しました。'}}"
if errorlevel 1 echo 検索または削除に失敗しました。使用中のファイルは削除できない場合があります。

pause
