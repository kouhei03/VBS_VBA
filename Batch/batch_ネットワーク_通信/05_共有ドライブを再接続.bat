@echo off
rem 概要: 「共有ドライブを再接続」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "05_共有ドライブを再接続.bat"
rem 実行結果（想定結果）: 「共有ドライブを再接続」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では本文内のWindowsまたはPowerShellコマンドを実行します。分岐がある場合は選択条件に従います。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set "DRIVE="
set "SHARE="
set /p "DRIVE=割り当てるドライブ文字（例: Z）: "
set /p "SHARE=共有フォルダの UNC パス（例: \\server\share）: "

if not "%DRIVE:~1,1%"=="" (
    echo ドライブ文字は1文字で指定してください。
    exit /b 1
)

if "%DRIVE%"=="" exit /b 1
net use "%DRIVE%:" "%SHARE%" /persistent:yes
if errorlevel 1 echo 接続できません。共有パス、VPN、アクセス権を確認してください。

pause
