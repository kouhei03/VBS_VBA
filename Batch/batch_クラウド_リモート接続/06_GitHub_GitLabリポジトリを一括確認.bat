@echo off
rem 概要: 「GitHub GitLabリポジトリを一括確認」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "06_GitHub_GitLabリポジトリを一括確認.bat"
rem 実行結果（想定結果）: 「GitHub GitLabリポジトリを一括確認」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: OpenSSHでリモート接続またはコマンド実行、リモートGitリポジトリをclone。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set "REPOS="
set "ROOT="
set /p "REPOS=GitHub/GitLab の HTTPS または SSH URL（複数は | 区切り）: "
set /p "ROOT=保存先フォルダ: "
if "%REPOS%"=="" exit /b 1
if not exist "%ROOT%\." mkdir "%ROOT%"

if errorlevel 1 (
    echo 保存先を作成できませんでした。
    exit /b 1
)

powershell -NoProfile -Command "$root=[IO.Path]::GetFullPath($env:ROOT); $urls=$env:REPOS -split '\|' | ForEach-Object {$_.Trim()} | Where-Object {$_}; foreach($url in $urls){if($url -notmatch '^(https://(github\.com|gitlab\.com)/|git@(github\.com|gitlab\.com):)[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+(\.git)?$'){Write-Error ('許可されていないURL形式: '+$url); exit 1}; $name=([IO.Path]::GetFileNameWithoutExtension(($url -split '[:/]')[-1])); $path=Join-Path $root $name; if(Test-Path (Join-Path $path '.git')){Write-Host ('fetch: '+$name); & git -C $path fetch --all --prune; if($LASTEXITCODE -ne 0){exit $LASTEXITCODE}}elseif(Test-Path $path){Write-Error ('同名の保存先が Git リポジトリではありません: '+$path); exit 1}else{Write-Host ('clone: '+$name); & git clone -- $url $path; if($LASTEXITCODE -ne 0){exit $LASTEXITCODE}}}"
if errorlevel 1 echo リポジトリの clone/fetch に失敗しました。Git の認証や保存先を確認してください。

pause
