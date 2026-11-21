@echo off
rem 概要: 「CSVをヘッダー付きで統合」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "11_CSVをヘッダー付きで統合.bat"
rem 実行結果（想定結果）: 「CSVをヘッダー付きで統合」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: Get-ChildItemで対象を列挙、Get-Contentでファイル内容を読み込み、Remove-Itemで対象を削除、PowerShellのImport-Csvで対象を処理。複数の分岐がある場合は条件に応じて実行します。
rem   3. 結果やログを本文で指定したファイルへ保存し、必要に応じて画面にも表示します。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set "SOURCE="
set "OUTPUT="
set /p "SOURCE=統合する CSV ファイルのフォルダ: "
set /p "OUTPUT=統合後の CSV ファイルパス: "

if not exist "%SOURCE%\." (
    echo 入力フォルダが見つかりません。

    pause
    exit /b 1
)

if not exist "%OUTPUT%" goto start_merge
set "CONFIRM="
set /p "CONFIRM=既存の出力ファイルを上書きする場合は OVERWRITE と入力: "
if not "%CONFIRM%"=="OVERWRITE" exit /b 0

:start_merge

powershell -NoProfile -Command "$ErrorActionPreference='Stop'; $s=[IO.Path]::GetFullPath($env:SOURCE).TrimEnd('\'); $o=[IO.Path]::GetFullPath($env:OUTPUT); if($o.StartsWith($s+'\',[StringComparison]::OrdinalIgnoreCase)){throw '出力 CSV は入力フォルダの外に指定してください'}; $files=@(Get-ChildItem -LiteralPath $s -Filter '*.csv' -File | Sort-Object Name); if(-not $files.Count){throw '統合対象 CSV がありません'}; $header=(Get-Content -LiteralPath $files[0].FullName -TotalCount 1); if(-not $header){throw 'CSV ヘッダーがありません'}; $temp=$o+'.tmp'; if(Test-Path -LiteralPath $temp){Remove-Item -LiteralPath $temp -Force}; foreach($f in $files){$h=Get-Content -LiteralPath $f.FullName -TotalCount 1; if($h -ne $header){throw ('ヘッダーが一致しません: '+$f.FullName)}; Import-Csv -LiteralPath $f.FullName | Export-Csv -LiteralPath $temp -NoTypeInformation -Encoding UTF8 -Append}; Move-Item -LiteralPath $temp -Destination $o -Force; Write-Host ('統合ファイル数: '+$files.Count)"
if errorlevel 1 echo CSV の統合に失敗しました。

pause
