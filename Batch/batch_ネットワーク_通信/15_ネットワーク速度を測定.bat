@echo off
rem 概要: 「ネットワーク速度を測定」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "15_ネットワーク速度を測定.bat"
rem 実行結果（想定結果）: 「ネットワーク速度を測定」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. スクリプト内の設定値を使い、実行条件を確認します。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: HTTPリクエストまたはファイル転送を実行。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
where curl.exe >nul 2>&1

if errorlevel 1 (
    echo curl.exe が見つかりません。Windows 10/11 の curl を使用してください。

    pause
    exit /b 1
)

echo Cloudflare から約10 MBをダウンロードして速度を測定します。
echo 外部サービスへの接続とデータ転送が発生します。

curl.exe -L --max-time 60 -o NUL -s -w "平均受信速度: %{speed_download} bytes/sec\n" "https://speed.cloudflare.com/__down?bytes=10000000"
if errorlevel 1 echo 測定に失敗しました。ネットワーク接続を確認してください。

pause
