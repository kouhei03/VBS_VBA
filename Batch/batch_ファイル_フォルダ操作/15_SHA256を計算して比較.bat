@echo off
rem 概要: 「SHA256を計算して比較」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "15_SHA256を計算して比較.bat"
rem 実行結果（想定結果）: 「SHA256を計算して比較」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: Get-FileHashでハッシュ値を計算。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set /p "FILE1=比較する1つ目のファイル: "
set /p "FILE2=比較する2つ目のファイル: "

if not exist "%FILE1%" (
    echo 1つ目のファイルが見つかりません。

    pause
    exit /b 1
)

if not exist "%FILE2%" (
    echo 2つ目のファイルが見つかりません。

    pause
    exit /b 1
)

powershell -NoProfile -Command "$ErrorActionPreference='Stop'; $a=Get-FileHash -LiteralPath $env:FILE1 -Algorithm SHA256; $b=Get-FileHash -LiteralPath $env:FILE2 -Algorithm SHA256; Write-Host ('1: '+$a.Hash+'  '+$a.Path); Write-Host ('2: '+$b.Hash+'  '+$b.Path); if($a.Hash -eq $b.Hash){Write-Host '一致しています。'}else{Write-Host '一致しません。'}"
if errorlevel 1 echo ハッシュを計算できませんでした。

pause
