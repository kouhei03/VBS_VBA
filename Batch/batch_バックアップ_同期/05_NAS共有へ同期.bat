@echo off
rem 概要: 「NAS共有へ同期」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "05_NAS共有へ同期.bat"
rem 実行結果（想定結果）: 「NAS共有へ同期」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: Robocopyでファイルをコピーまたは同期。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set /p "SOURCE=同期元フォルダ: "
set /p "NAS=NAS共有フォルダ（例: \\NAS\Share\Backup）: "

if not exist "%SOURCE%\." (
    echo 同期元が見つかりません。

    pause
    exit /b 1
)

if not exist "%NAS%\." (
    echo NAS共有に接続できません。ネットワークとアクセス権を確認してください。

    pause
    exit /b 1
)

echo 1回分の追加・更新同期を実行します。定期実行はタスク スケジューラに登録してください。

robocopy "%SOURCE%" "%NAS%" /E /Z /FFT /R:3 /W:5 /XJ
if errorlevel 8 echo NAS 同期に失敗しました。終了コード: %ERRORLEVEL%

pause
