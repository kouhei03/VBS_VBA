@echo off
rem 概要: 「古い世代を整理」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "11_古い世代を整理.bat"
rem 実行結果（想定結果）: 「古い世代を整理」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: Get-ChildItemで対象を列挙、Remove-Itemで対象を削除。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set /p "ROOT=世代別バックアップの親フォルダ: "
set /p "KEEP=残す世代数（例: 5）: "

if not exist "%ROOT%\." (
    echo 対象フォルダが見つかりません。

    pause
    exit /b 1
)

powershell -NoProfile -Command "$ErrorActionPreference='Stop'; $keep=0; if(-not [int]::TryParse($env:KEEP,[ref]$keep) -or $keep -lt 1){throw '残す世代数は1以上の整数にしてください'}; $dirs=@(Get-ChildItem -LiteralPath $env:ROOT -Directory -Force | Sort-Object LastWriteTime -Descending); $old=@($dirs | Select-Object -Skip $keep); $old | Select-Object -ExpandProperty FullName; if($old.Count -eq 0){Write-Host '削除対象はありません。'}else{$answer=Read-Host '上記フォルダを完全削除する場合は DELETE と入力してください'; if($answer -ceq 'DELETE'){foreach($d in $old){Remove-Item -LiteralPath $d.FullName -Recurse -Force}; Write-Host ('削除世代数: '+$old.Count)}else{Write-Host '中止しました。'}}"
if errorlevel 1 echo 世代整理に失敗しました。

pause
