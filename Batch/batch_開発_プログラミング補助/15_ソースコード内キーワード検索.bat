@echo off
rem 概要: 「ソースコード内キーワード検索」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "15_ソースコード内キーワード検索.bat"
rem 実行結果（想定結果）: 「ソースコード内キーワード検索」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: Get-ChildItemで対象を列挙、Gitでリポジトリを操作、Select-Stringで指定キーワードを検索。複数の分岐がある場合は条件に応じて実行します。
rem   3. 結果やログを本文で指定したファイルへ保存し、必要に応じて画面にも表示します。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set "ROOT="
set "KEYWORD="
set /p "ROOT=検索するソースフォルダ: "
set /p "KEYWORD=検索キーワード: "

if not exist "%ROOT%\." (
    echo フォルダが見つかりません。

    pause
    exit /b 1
)

if "%KEYWORD%"=="" exit /b 1
set "OUTPUT=%~dp0keyword_search.txt"

powershell -NoProfile -Command "$ErrorActionPreference='Stop'; $files=@(Get-ChildItem -LiteralPath $env:ROOT -Recurse -File | Where-Object {$_.FullName -notmatch '[\\/](\.git|node_modules|bin|obj)[\\/]' -and $_.Extension -in '.py','.js','.jsx','.ts','.tsx','.cs','.java','.cpp','.h','.go','.rs','.html','.css','.json','.md'}); $matches=@(); if($files.Count){$matches=@(Select-String -Path $files.FullName -Pattern $env:KEYWORD -SimpleMatch -ErrorAction SilentlyContinue)}; $matches | Out-File -LiteralPath $env:OUTPUT -Encoding utf8"
if errorlevel 1 echo 検索またはログ出力に失敗しました。
echo 結果: %~dp0keyword_search.txt

pause
