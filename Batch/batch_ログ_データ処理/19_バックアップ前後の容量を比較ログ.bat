@echo off
rem 概要: 「バックアップ前後の容量を比較ログ」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "19_バックアップ前後の容量を比較ログ.bat"
rem 実行結果（想定結果）: 「バックアップ前後の容量を比較ログ」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: Get-ChildItemで対象を列挙、Robocopyでファイルをコピーまたは同期、Add-Contentでログへ追記。複数の分岐がある場合は条件に応じて実行します。
rem   3. 結果やログを本文で指定したファイルへ保存し、必要に応じて画面にも表示します。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set "SOURCE="
set "DEST="
set "LOG=%~dp0backup_comparison.log"
set /p "SOURCE=バックアップ元フォルダ: "
set /p "DEST=バックアップ先フォルダ: "

if not exist "%SOURCE%\." (
    echo バックアップ元が見つかりません。

    pause
    exit /b 1
)

if not exist "%DEST%\." mkdir "%DEST%"

powershell -NoProfile -Command "$ErrorActionPreference='Stop'; $s=[IO.Path]::GetFullPath($env:SOURCE).TrimEnd('\'); $d=[IO.Path]::GetFullPath($env:DEST).TrimEnd('\'); if($d -eq $s -or $d.StartsWith($s+'\',[StringComparison]::OrdinalIgnoreCase)){throw 'バックアップ先を元フォルダ内に指定できません'}; function Stats {param($p); $f=@(Get-ChildItem -LiteralPath $p -File -Recurse -Force); $sum=($f | Measure-Object Length -Sum).Sum; if($null -eq $sum){$sum=0}; [pscustomobject]@{Count=$f.Count; Bytes=[int64]$sum}}; $src=Stats $s; $before=Stats $d; & robocopy $s $d /E /R:2 /W:2 /XJ; $rc=$LASTEXITCODE; $after=Stats $d; $line='{0} Source={1} files/{2} bytes; DestBefore={3}/{4}; DestAfter={5}/{6}; Robocopy={7}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$src.Count,$src.Bytes,$before.Count,$before.Bytes,$after.Count,$after.Bytes,$rc; Add-Content -LiteralPath $env:LOG -Value $line -Encoding UTF8; Write-Host $line; if($rc -ge 8){exit $rc}"
if errorlevel 1 echo バックアップに失敗しました。詳細はログを確認してください。

pause
