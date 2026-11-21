@echo off
rem 概要: 「IPアドレスとMACアドレスを保存」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "04_IPアドレスとMACアドレスを保存.bat"
rem 実行結果（想定結果）: 「IPアドレスとMACアドレスを保存」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. スクリプト内の設定値を使い、実行条件を確認します。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: PowerShellのGet-NetIPConfigurationで対象を処理、PowerShellのGet-NetAdapterで対象を処理。複数の分岐がある場合は条件に応じて実行します。
rem   3. 結果やログを本文で指定したファイルへ保存し、必要に応じて画面にも表示します。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set "OUTPUT=%~dp0network_info.txt"

powershell -NoProfile -Command "$ErrorActionPreference='Stop'; '取得日時: '+(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | Out-File -LiteralPath $env:OUTPUT -Encoding utf8; Get-NetIPConfiguration | Format-List InterfaceAlias,InterfaceDescription,IPv4Address,IPv6Address,IPv4DefaultGateway,IPv6DefaultGateway,DNSServer | Out-File -LiteralPath $env:OUTPUT -Append -Encoding utf8; Get-NetAdapter | Format-Table Name,Status,MacAddress,LinkSpeed -AutoSize | Out-File -LiteralPath $env:OUTPUT -Append -Encoding utf8"

if errorlevel 1 (
    echo ネットワーク情報を保存できませんでした。

    pause
    exit /b 1
)

echo 保存しました: %OUTPUT%

pause
