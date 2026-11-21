@echo off
rem 概要: 「依存パッケージを更新」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "16_依存パッケージを更新.bat"
rem 実行結果（想定結果）: 「依存パッケージを更新」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では本文内のWindowsまたはPowerShellコマンドを実行します。分岐がある場合は選択条件に従います。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
echo 1: npm update（package.json の範囲内で更新）
echo 2: pip の requirements.txt を更新
set "TYPE="
set /p "TYPE=パッケージ管理ツールを選択してください (1/2): "
set "PROJECT="
set /p "PROJECT=プロジェクトフォルダ: "
set "PYTHON=python"
if "%TYPE%"=="2" set /p "PYTHON=使用する python.exe（仮想環境を推奨。空欄なら python）: "

if not exist "%PROJECT%\." (
    echo フォルダが見つかりません。

    pause
    exit /b 1
)

echo 更新により互換性が変わる場合があります。先に変更をコミットしてください。
set "CONFIRM="
set /p "CONFIRM=続行する場合は UPDATE と入力してください: "
if not "%CONFIRM%"=="UPDATE" exit /b 0

if "%TYPE%"=="1" (

    if not exist "%PROJECT%\package.json" (
        echo package.json が見つかりません。
        exit /b 1
    )

    cd /d "%PROJECT%"
    npm update
) else if "%TYPE%"=="2" (

    if not exist "%PROJECT%\requirements.txt" (
        echo requirements.txt が見つかりません。
        exit /b 1
    )

    cd /d "%PROJECT%"
    if "%PYTHON%"=="" set "PYTHON=python"
    "%PYTHON%" -m pip install --upgrade -r requirements.txt
) else (
    echo ツールが選択されていません。
    exit /b 1
)

if errorlevel 1 echo 依存関係の更新に失敗しました。

pause
