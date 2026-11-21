@echo off
rem 概要: 「文字コードを一括変換」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "04_文字コードを一括変換.bat"
rem 実行結果（想定結果）: 「文字コードを一括変換」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: Get-ChildItemで対象を列挙、PowerShellのNew-Itemで対象を処理。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set "SOURCE="
set "DEST="
set "FROM="
set "TO="
set /p "SOURCE=変換するテキストファイルの親フォルダ: "
set /p "DEST=出力先フォルダ（原本は変更しません）: "
echo 入力文字コード: 1=UTF-8 / 2=Shift-JIS
set /p "FROM=選択してください (1/2): "
echo 出力文字コード: 1=UTF-8 / 2=Shift-JIS
set /p "TO=選択してください (1/2): "

if not exist "%SOURCE%\." (
    echo 入力フォルダが見つかりません。

    pause
    exit /b 1
)

if not exist "%DEST%\." mkdir "%DEST%"

powershell -NoProfile -Command "$ErrorActionPreference='Stop'; $s=[IO.Path]::GetFullPath($env:SOURCE).TrimEnd('\'); $d=[IO.Path]::GetFullPath($env:DEST).TrimEnd('\'); if($d -eq $s -or $d.StartsWith($s+'\',[StringComparison]::OrdinalIgnoreCase)){throw '出力先は入力フォルダの外に指定してください'}; if($env:FROM -notin '1','2' -or $env:TO -notin '1','2'){throw '文字コードの選択が不正です'}; $from=if($env:FROM -eq '1'){[Text.Encoding]::UTF8}else{[Text.Encoding]::GetEncoding(932)}; $to=if($env:TO -eq '1'){[Text.Encoding]::UTF8}else{[Text.Encoding]::GetEncoding(932)}; $files=@(Get-ChildItem -LiteralPath $s -Recurse -File -Include *.txt,*.csv,*.log); foreach($f in $files){$rel=$f.FullName.Substring($s.Length).TrimStart('\'); $out=Join-Path $d $rel; New-Item -ItemType Directory -Path (Split-Path -Parent $out) -Force | Out-Null; $text=[IO.File]::ReadAllText($f.FullName,$from); [IO.File]::WriteAllText($out,$text,$to)}; Write-Host ('変換数: '+$files.Count)"
if errorlevel 1 echo 文字コード変換に失敗しました。

pause
