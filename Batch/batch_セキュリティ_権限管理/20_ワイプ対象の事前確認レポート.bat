@echo off
rem 概要: 「ワイプ対象の事前確認レポート」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "20_ワイプ対象の事前確認レポート.bat"
rem 実行結果（想定結果）: 「ワイプ対象の事前確認レポート」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: Get-ChildItemで対象を列挙、Set-Contentで結果をファイルへ書き込み。複数の分岐がある場合は条件に応じて実行します。
rem   3. 結果やログを本文で指定したファイルへ保存し、必要に応じて画面にも表示します。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set "TARGET="
set "OUTPUT=%~dp0wipe_dry_run.txt"
set /p "TARGET=確認するデータフォルダ（システムフォルダ不可）: "

if not exist "%TARGET%\." (
    echo フォルダが見つかりません。

    pause
    exit /b 1
)

echo このスクリプトは削除・上書きを行わず、対象の概要だけを記録します。
set "CONFIRM="
set /p "CONFIRM=事前確認レポートを作成する場合は REVIEW と入力してください: "
if not "%CONFIRM%"=="REVIEW" exit /b 0

powershell -NoProfile -Command "$p=[IO.Path]::GetFullPath($env:TARGET).TrimEnd('\'); $protected=@($env:WINDIR,$env:ProgramFiles,${env:ProgramFiles(x86)},$env:SystemDrive); foreach($root in $protected){if($root -and $p.Equals([IO.Path]::GetFullPath($root).TrimEnd('\'),[StringComparison]::OrdinalIgnoreCase)){throw 'システム領域全体は対象にできません'}}; $files=@(Get-ChildItem -LiteralPath $p -File -Recurse -Force); $sum=($files | Measure-Object Length -Sum).Sum; if($null -eq $sum){$sum=0}; @('DRY RUN ONLY - no files deleted',('Target: '+$p),('Generated: '+(Get-Date)),('File count: '+$files.Count),('Total bytes: '+$sum)) | Set-Content -LiteralPath $env:OUTPUT -Encoding UTF8; Write-Host ('ファイル数: '+$files.Count+' / 合計バイト数: '+$sum)"
if errorlevel 1 echo レポートを作成できませんでした。
echo レポート: %OUTPUT%

pause
