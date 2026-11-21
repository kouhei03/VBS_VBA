@echo off
rem 概要: 「S3古いバックアップを整理」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "11_S3古いバックアップを整理.bat"
rem 実行結果（想定結果）: 「S3古いバックアップを整理」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: S3の指定プレフィックス内を一覧、PowerShellのConvertFrom-Jsonで対象を処理、対象S3オブジェクトを削除。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set "BUCKET="
set "PREFIX="
set "DAYS="
set "REGION="
set /p "BUCKET=S3 バケット名: "
set /p "PREFIX=対象プレフィックス（例: backups/project/）: "
set /p "DAYS=何日より古いオブジェクトを対象にするか: "
set /p "REGION=AWS リージョン: "

powershell -NoProfile -Command "if($env:BUCKET -notmatch '^[A-Za-z0-9.-]{3,63}$' -or $env:DAYS -notmatch '^\d{1,5}$' -or [int]$env:DAYS -lt 1 -or [int]$env:DAYS -gt 36500){throw 'バケット名または日数が不正です'}"
if errorlevel 1 exit /b 1
echo 対象を一覧表示し、確認後に指定プレフィックス内の古いオブジェクトだけを削除します。
set "CONFIRM="
set /p "CONFIRM=一覧を取得する場合は REVIEW と入力: "
if not "%CONFIRM%"=="REVIEW" exit /b 0

powershell -NoProfile -Command "$json=& aws s3api list-objects-v2 --bucket $env:BUCKET --prefix $env:PREFIX --region $env:REGION --output json; if($LASTEXITCODE -ne 0){exit $LASTEXITCODE}; $data=$json | ConvertFrom-Json; $cutoff=[DateTime]::UtcNow.AddDays(-[int]$env:DAYS); $items=@($data.Contents | Where-Object {[DateTime]::Parse($_.LastModified).ToUniversalTime() -lt $cutoff}); if($items.Count -eq 0){Write-Host '削除対象はありません'; exit 0}; if($items.Count -gt 500){throw '対象が500件を超えています。プレフィックスを狭めてください'}; $items | Select-Object Key,LastModified,Size | Format-Table -AutoSize; $confirm=Read-Host '表示された対象だけを削除する場合は DELETE と入力'; if($confirm -cne 'DELETE'){Write-Host 'キャンセルしました'; exit 0}; foreach($item in $items){& aws s3api delete-object --bucket $env:BUCKET --key $item.Key --region $env:REGION | Out-Null; if($LASTEXITCODE -ne 0){Write-Error ('削除失敗: '+$item.Key); exit 1}}; Write-Host ('削除件数: '+$items.Count)"
if errorlevel 1 echo 一覧取得または削除に失敗しました。AWS CLI の権限とバケットを確認してください。

pause
