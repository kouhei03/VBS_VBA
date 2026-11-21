@echo off
rem 概要: 「プロセス一覧をテキスト保存」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "16_プロセス一覧をテキスト保存.bat"
rem 実行結果（想定結果）: 「プロセス一覧をテキスト保存」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. スクリプト内の設定値を使い、実行条件を確認します。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: Get-Processで対象プロセスを取得。複数の分岐がある場合は条件に応じて実行します。
rem   3. 結果やログを本文で指定したファイルへ保存し、必要に応じて画面にも表示します。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set "OUTPUT=%~dp0process_list.txt"

powershell -NoProfile -Command "Get-Process | Sort-Object ProcessName | Select-Object ProcessName,Id,CPU,WorkingSet64,Path | Format-Table -AutoSize | Out-File -LiteralPath $env:OUTPUT -Encoding UTF8"

if errorlevel 1 (
    echo プロセス一覧を保存できませんでした。

    pause
    exit /b 1
)

echo 保存しました: %OUTPUT%

pause
