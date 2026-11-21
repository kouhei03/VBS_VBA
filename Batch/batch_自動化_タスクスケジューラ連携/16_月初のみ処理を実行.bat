@echo off
rem 概要: 「月初のみ処理を実行」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "16_月初のみ処理を実行.bat"
rem 実行結果（想定結果）: 「月初のみ処理を実行」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. スクリプト内の設定値を使い、実行条件を確認します。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では本文内のWindowsまたはPowerShellコマンドを実行します。分岐がある場合は選択条件に従います。
rem   3. 結果やログを本文で指定したファイルへ保存し、必要に応じて画面にも表示します。

chcp 65001 >nul
setlocal
for /f "delims=" %%D in ('powershell -NoProfile -Command "(Get-Date).Day"') do set "DAY=%%D"

if not "%DAY%"=="1" (
    echo 今日は月初ではありません。処理を終了します。
    exit /b 0
)

echo 月初を検出しました。月初専用処理を実行します。
echo %date% %time% - first day of month>>"%~dp0month_start.log"
echo 月初処理をここに追加してください。
