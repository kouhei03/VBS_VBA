@echo off
rem 概要: 「件数ごとに別フォルダへ分割移動」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "17_件数ごとに別フォルダへ分割移動.bat"
rem 実行結果（想定結果）: 「件数ごとに別フォルダへ分割移動」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: Get-ChildItemで対象を列挙、PowerShellのNew-Itemで対象を処理、Move-Itemでファイルを移動、ファイルを移動。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set /p "SOURCE=移動元フォルダ（直下のファイルのみ）: "
set /p "DEST=分割先の親フォルダ: "
set /p "COUNT=1フォルダあたりのファイル数: "

if not exist "%SOURCE%\." (
    echo 移動元フォルダが見つかりません。

    pause
    exit /b 1
)

if not exist "%DEST%\." mkdir "%DEST%"
echo ファイル名順に並べ、指定件数ごとに移動します。
set /p "CONFIRM=続行しますか？ (Y/N): "
if /i not "%CONFIRM%"=="Y" exit /b 0

powershell -NoProfile -Command "$ErrorActionPreference='Stop'; $n=0; if(-not [int]::TryParse($env:COUNT,[ref]$n) -or $n -le 0){throw '件数は1以上の整数を指定してください'}; $files=@(Get-ChildItem -LiteralPath $env:SOURCE -File | Sort-Object Name); $plan=@(for($i=0;$i -lt $files.Count;$i++){$group=[math]::Floor($i/$n)+1; $folder=Join-Path $env:DEST ('分割_{0:D3}' -f $group); $target=Join-Path $folder $files[$i].Name; if(Test-Path -LiteralPath $target){throw ('移動先に同名ファイルがあります: '+$target)}; [pscustomobject]@{File=$files[$i];Folder=$folder}}); foreach($item in $plan){New-Item -ItemType Directory -Path $item.Folder -Force | Out-Null; Move-Item -LiteralPath $item.File.FullName -Destination $item.Folder}; Write-Host ('移動数: '+$plan.Count)"
if errorlevel 1 echo 分割移動に失敗しました。

pause
