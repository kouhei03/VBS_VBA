@echo off
rem 概要: 「AWS SAMローカルデプロイテスト」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "08_AWS_SAMローカルデプロイテスト.bat"
rem 実行結果（想定結果）: 「AWS SAMローカルデプロイテスト」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: SAMテンプレートを検証、SAMアプリケーションをビルド、SAM関数をローカル実行。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set "PROJECT="
set "FUNCTION="
set /p "PROJECT=SAM テンプレートのあるプロジェクトフォルダ: "

if not exist "%PROJECT%\template.yaml" if not exist "%PROJECT%\template.yml" (
    echo template.yaml または template.yml が見つかりません。

    pause
    exit /b 1
)

set /p "FUNCTION=ローカル実行する関数の論理ID（空欄なら指定なし）: "
pushd "%PROJECT%"

sam validate
if errorlevel 1 goto failed

sam build
if errorlevel 1 goto failed

if "%FUNCTION%"=="" (
    echo ビルドが完了しました。関数を実行する場合は sam local invoke を利用してください。
    popd

    pause
    exit /b 0
)

sam local invoke "%FUNCTION%"
if errorlevel 1 goto failed
popd

pause
exit /b 0

:failed

set "RESULT=%ERRORLEVEL%"
popd
echo SAM の検証、ビルド、またはローカル実行に失敗しました。

pause
exit /b %RESULT%
