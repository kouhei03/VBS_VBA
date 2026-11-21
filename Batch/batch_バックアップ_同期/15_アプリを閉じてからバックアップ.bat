@echo off
rem 概要: 「アプリを閉じてからバックアップ」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "15_アプリを閉じてからバックアップ.bat"
rem 実行結果（想定結果）: 「アプリを閉じてからバックアップ」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: Get-Processで対象プロセスを取得、Robocopyでファイルをコピーまたは同期。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set /p "PROCESS=終了するアプリのプロセス名（.exe なし）: "
set /p "SOURCE=バックアップ元フォルダ: "
set /p "DEST=バックアップ先フォルダ: "

if not exist "%SOURCE%\." (
    echo バックアップ元が見つかりません。

    pause
    exit /b 1
)

echo アプリに通常終了を要求します。未保存の作業は事前に保存してください。

powershell -NoProfile -Command "$ErrorActionPreference='Stop'; $p=@(Get-Process -Name $env:PROCESS -ErrorAction SilentlyContinue); if($p.Count -eq 0){Write-Host '対象アプリは起動していません。'}else{foreach($x in $p){if($x.MainWindowHandle -eq 0){throw '画面のないプロセスがあるため安全のため中止します'}; [void]$x.CloseMainWindow()}; Start-Sleep -Seconds 10; if(Get-Process -Name $env:PROCESS -ErrorAction SilentlyContinue){throw 'アプリが終了しませんでした。強制終了せずバックアップを中止します'}; Write-Host 'アプリを終了しました。'}"

if errorlevel 1 (
    echo バックアップを中止しました。

    pause
    exit /b 1
)

robocopy "%SOURCE%" "%DEST%" /E /Z /R:3 /W:5 /XJ
if errorlevel 8 echo バックアップに失敗しました。終了コード: %ERRORLEVEL%

pause
