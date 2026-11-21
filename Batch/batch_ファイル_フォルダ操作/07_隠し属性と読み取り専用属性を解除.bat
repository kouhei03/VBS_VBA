@echo off
rem 概要: 「隠し属性と読み取り専用属性を解除」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "07_隠し属性と読み取り専用属性を解除.bat"
rem 実行結果（想定結果）: 「隠し属性と読み取り専用属性を解除」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: Get-ChildItemで対象を列挙。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set /p "TARGET_DIR=対象フォルダ（サブフォルダも対象）: "

if not exist "%TARGET_DIR%\." (
    echo 対象フォルダが見つかりません。

    pause
    exit /b 1
)

echo ファイルの隠し属性と読み取り専用属性を解除します。
set /p "CONFIRM=続行しますか？ (Y/N): "
if /i not "%CONFIRM%"=="Y" exit /b 0

powershell -NoProfile -Command "$ErrorActionPreference='Stop'; Get-ChildItem -LiteralPath $env:TARGET_DIR -File -Recurse -Force | ForEach-Object { $_.Attributes = $_.Attributes -band (-bnot ([IO.FileAttributes]::Hidden -bor [IO.FileAttributes]::ReadOnly)) }"
if errorlevel 1 echo 属性を変更できないファイルがありました。

pause
