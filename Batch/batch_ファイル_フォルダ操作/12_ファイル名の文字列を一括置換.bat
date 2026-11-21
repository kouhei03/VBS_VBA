@echo off
rem 概要: 「ファイル名の文字列を一括置換」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "12_ファイル名の文字列を一括置換.bat"
rem 実行結果（想定結果）: 「ファイル名の文字列を一括置換」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: Get-ChildItemで対象を列挙、PowerShellのRename-Itemで対象を処理。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set /p "TARGET_DIR=対象フォルダ（直下のファイルのみ）: "
set /p "OLD_TEXT=ファイル名内の置換対象文字列: "
set /p "NEW_TEXT=置換後の文字列（空欄で削除）: "

if not exist "%TARGET_DIR%\." (
    echo 対象フォルダが見つかりません。

    pause
    exit /b 1
)

if "%OLD_TEXT%"=="" (
    echo 置換対象文字列を空欄にはできません。

    pause
    exit /b 1
)

echo ファイル名に含まれる文字列を置換します。
set /p "CONFIRM=続行しますか？ (Y/N): "
if /i not "%CONFIRM%"=="Y" exit /b 0

powershell -NoProfile -Command "$ErrorActionPreference='Stop'; $files=@(Get-ChildItem -LiteralPath $env:TARGET_DIR -File -Force); $plan=@(foreach($f in $files){$new=$f.Name.Replace($env:OLD_TEXT,$env:NEW_TEXT); if($new -ne $f.Name){$destination=Join-Path $f.DirectoryName $new; if(Test-Path -LiteralPath $destination){throw ('変更先が既にあります: '+$new)}; [pscustomobject]@{File=$f;Name=$new}}}); $names=@($plan | ForEach-Object Name); if(($names | Select-Object -Unique).Count -ne $names.Count){throw '変更後のファイル名が重複します'}; foreach($item in $plan){Rename-Item -LiteralPath $item.File.FullName -NewName $item.Name}"
if errorlevel 1 echo 名前を変更できませんでした。変更先の重複などを確認してください。

pause
