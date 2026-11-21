@echo off
rem 概要: 「今日の日付のスナップショット」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "10_今日の日付のスナップショット.bat"
rem 実行結果（想定結果）: 「今日の日付のスナップショット」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: Robocopyでファイルをコピーまたは同期、ファイルをコピー。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set /p "SOURCE=スナップショット元フォルダ: "
set /p "PARENT=スナップショット保存先の親フォルダ: "

if not exist "%SOURCE%\." (
    echo 対象フォルダが見つかりません。

    pause
    exit /b 1
)

for /f %%D in ('powershell -NoProfile -Command "Get-Date -Format yyyy-MM-dd"') do set "TODAY=%%D"
set "DEST=%PARENT%\Snapshot_%TODAY%"
if not exist "%PARENT%\." mkdir "%PARENT%"

if exist "%DEST%\." (
    echo 今日のスナップショットは既にあります。既存内容へ追加・更新します。
)

robocopy "%SOURCE%" "%DEST%" /E /COPY:DAT /DCOPY:DAT /R:2 /W:2 /XJ
if errorlevel 8 echo スナップショット作成に失敗しました。終了コード: %ERRORLEVEL%

pause
