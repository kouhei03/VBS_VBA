@echo off
rem 概要: 「Downloadsを種類別に振り分け」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "13_Downloadsを種類別に振り分け.bat"
rem 実行結果（想定結果）: 「Downloadsを種類別に振り分け」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: PowerShellのNew-Itemで対象を処理、Get-ChildItemで対象を列挙、Move-Itemでファイルを移動、ファイルを移動。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set "DOWNLOADS=%USERPROFILE%\Downloads"

if not exist "%DOWNLOADS%\." (
    echo Downloads フォルダが見つかりません: %DOWNLOADS%

    pause
    exit /b 1
)

echo 画像・動画・ドキュメントを Downloads 内の各フォルダへ移動します。
set /p "CONFIRM=続行しますか？ (Y/N): "
if /i not "%CONFIRM%"=="Y" exit /b 0

powershell -NoProfile -Command "$ErrorActionPreference='Stop'; $root=$env:DOWNLOADS; $groups=@{Images=@('.jpg','.jpeg','.png','.gif','.bmp','.webp','.tif','.tiff'); Videos=@('.mp4','.mov','.avi','.mkv','.wmv','.webm','.m4v'); Documents=@('.pdf','.doc','.docx','.xls','.xlsx','.ppt','.pptx','.txt','.rtf','.csv')}; foreach($name in $groups.Keys){$dest=Join-Path $root $name; New-Item -ItemType Directory -Path $dest -Force | Out-Null; Get-ChildItem -LiteralPath $root -File | Where-Object {$groups[$name] -contains $_.Extension.ToLowerInvariant()} | Move-Item -Destination $dest -ErrorAction Continue}"
if errorlevel 1 echo 一部のファイルを移動できませんでした。

pause
