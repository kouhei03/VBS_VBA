@echo off
rem 概要: 「USB接続を検知してバックアップ」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "02_USB接続を検知してバックアップ.bat"
rem 実行結果（想定結果）: 「USB接続を検知してバックアップ」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: PowerShellのGet-Diskで対象を処理、PowerShellのGet-Partitionで対象を処理、Get-Volumeでボリューム情報を取得、Robocopyでファイルをコピーまたは同期。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set /p "SOURCE=バックアップ元フォルダ: "

if not exist "%SOURCE%\." (
    echo バックアップ元が見つかりません。

    pause
    exit /b 1
)

echo USBドライブの接続を監視します。終了は Ctrl+C です。
echo USB接続の外付けHDDとリムーバブルドライブを検知し、PC名のフォルダへ追加コピーします。

powershell -NoProfile -Command "$known=@(); while($true){$usb=@(foreach($disk in (Get-Disk | Where-Object BusType -eq 'USB')){Get-Partition -DiskNumber $disk.Number | Where-Object DriveLetter | ForEach-Object {$_.DriveLetter+':'}}); $rem=@(Get-Volume | Where-Object {$_.DriveType -eq 'Removable' -and $_.DriveLetter} | ForEach-Object {$_.DriveLetter+':'}); $current=@($usb+$rem | Sort-Object -Unique); $new=@($current | Where-Object {$_ -notin $known}); foreach($drive in $new){$dest=Join-Path ($drive+'\') ('Backup_'+$env:COMPUTERNAME); $src=[IO.Path]::GetFullPath($env:SOURCE).TrimEnd('\'); $target=[IO.Path]::GetFullPath($dest).TrimEnd('\'); if($target.StartsWith($src+'\',[StringComparison]::OrdinalIgnoreCase) -or $target -eq $src){Write-Host ('コピー先が元フォルダ内のためスキップ: '+$dest); continue}; Write-Host ('接続を検知: '+$drive+' -> '+$dest); & robocopy $env:SOURCE $dest /E /R:2 /W:2 /XJ; if($LASTEXITCODE -ge 8){Write-Host ('バックアップ失敗: '+$LASTEXITCODE)}else{Write-Host 'バックアップ完了'}}; $known=$current; Start-Sleep -Seconds 5}"
