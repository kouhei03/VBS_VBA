@echo off
rem 概要: 「階層を保って特定ファイルをコピー」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "05_階層を保って特定ファイルをコピー.bat"
rem 実行結果（想定結果）: 「階層を保って特定ファイルをコピー」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: Get-ChildItemで対象を列挙、PowerShellのNew-Itemで対象を処理、Copy-Itemでファイルをコピー、ファイルをコピー。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set /p "SOURCE=検索元フォルダ: "
set /p "DEST=コピー先フォルダ: "
set /p "PATTERN=対象ファイル名またはワイルドカード（例: *.pdf）: "

if not exist "%SOURCE%\." (
    echo 検索元フォルダが見つかりません。

    pause
    exit /b 1
)

if not exist "%DEST%\." mkdir "%DEST%"
echo サブフォルダを含めて対象ファイルを検索し、階層を保ってコピーします。
set /p "CONFIRM=続行しますか？ (Y/N): "
if /i not "%CONFIRM%"=="Y" exit /b 0

powershell -NoProfile -Command "$ErrorActionPreference='Stop'; $src=(Resolve-Path -LiteralPath $env:SOURCE).Path.TrimEnd('\'); $dst=(Resolve-Path -LiteralPath $env:DEST).Path; if($dst.Equals($src,[StringComparison]::OrdinalIgnoreCase) -or $dst.StartsWith($src+'\',[StringComparison]::OrdinalIgnoreCase)){throw 'コピー先は検索元フォルダの外に指定してください'}; $items=@(Get-ChildItem -LiteralPath $src -File -Recurse -Filter $env:PATTERN); foreach($f in $items){$rel=$f.FullName.Substring($src.Length).TrimStart('\'); $out=Join-Path $dst $rel; if(Test-Path -LiteralPath $out){throw ('コピー先に同名ファイルがあります: '+$out)}; $parent=Split-Path -Parent $out; New-Item -ItemType Directory -Path $parent -Force | Out-Null; Copy-Item -LiteralPath $f.FullName -Destination $out}; Write-Host ('コピー数: ' + $items.Count)"
if errorlevel 1 echo コピーに失敗しました。

pause
