@echo off
rem 概要: 「フォルダサイズとファイル数を集計」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "19_フォルダサイズとファイル数を集計.bat"
rem 実行結果（想定結果）: 「フォルダサイズとファイル数を集計」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: Get-ChildItemで対象を列挙。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set /p "TARGET_DIR=集計するフォルダ: "

if not exist "%TARGET_DIR%\." (
    echo 対象フォルダが見つかりません。

    pause
    exit /b 1
)

powershell -NoProfile -Command "$ErrorActionPreference='Stop'; $files=@(Get-ChildItem -LiteralPath $env:TARGET_DIR -File -Recurse -Force); $bytes=($files | Measure-Object -Property Length -Sum).Sum; if($null -eq $bytes){$bytes=0}; Write-Host ('ファイル数（サブフォルダを含む）: '+$files.Count); Write-Host ('合計サイズ: '+$bytes+' bytes ('+[math]::Round($bytes/1MB,2)+' MB)')"
if errorlevel 1 echo 集計に失敗しました。

pause
