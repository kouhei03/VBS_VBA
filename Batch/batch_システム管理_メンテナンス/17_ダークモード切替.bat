@echo off
rem 概要: 「ダークモード切替」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "17_ダークモード切替.bat"
rem 実行結果（想定結果）: 「ダークモード切替」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: PowerShellのNew-Itemで対象を処理、PowerShellのNew-ItemPropertyで対象を処理。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
echo 1: ダークモード
echo 2: ライトモード
set "ACTION="
set /p "ACTION=テーマを選択してください (1/2): "

if "%ACTION%"=="1" (
    set "THEME=0"
) else if "%ACTION%"=="2" (
    set "THEME=1"
) else (
    echo 操作が選択されていません。
    exit /b 1
)

powershell -NoProfile -Command "$value=[int]$env:THEME; $path='HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize'; New-Item -Path $path -Force | Out-Null; New-ItemProperty -Path $path -Name AppsUseLightTheme -PropertyType DWord -Value $value -Force | Out-Null; New-ItemProperty -Path $path -Name SystemUsesLightTheme -PropertyType DWord -Value $value -Force | Out-Null"
if errorlevel 1 echo テーマを変更できませんでした。
echo 一部のアプリは再起動後にテーマが反映されます。

pause
