@echo off
rem 概要: 「祝日以外にバックアップ」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "19_祝日以外にバックアップ.bat"
rem 実行結果（想定結果）: 「祝日以外にバックアップ」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: Get-Contentでファイル内容を読み込み、Robocopyでファイルをコピーまたは同期。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set /p "HOLIDAYS=祝日一覧ファイル（yyyy-MM-dd を1行ずつ記載）: "
set /p "SOURCE=バックアップ元フォルダ: "
set /p "DEST=バックアップ先フォルダ: "

if not exist "%HOLIDAYS%" (
    echo 祝日一覧ファイルが見つかりません。

    pause
    exit /b 1
)

if not exist "%SOURCE%\." (
    echo バックアップ元が見つかりません。

    pause
    exit /b 1
)

powershell -NoProfile -Command "$today=Get-Date -Format 'yyyy-MM-dd'; $dates=Get-Content -LiteralPath $env:HOLIDAYS | ForEach-Object {$_.Trim()}; if($dates -contains $today){Write-Host ('今日は祝日 ('+$today+') のためバックアップをスキップします'); exit 0}; & robocopy $env:SOURCE $env:DEST /E /Z /R:3 /W:5 /XJ; if($LASTEXITCODE -ge 8){exit $LASTEXITCODE}"
if errorlevel 1 echo バックアップに失敗しました。

pause
