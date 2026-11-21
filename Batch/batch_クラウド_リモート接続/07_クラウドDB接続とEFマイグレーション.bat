@echo off
rem 概要: 「クラウドDB接続とEFマイグレーション」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "07_クラウドDB接続とEFマイグレーション.bat"
rem 実行結果（想定結果）: 「クラウドDB接続とEFマイグレーション」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: Test-NetConnectionで指定ポートへの接続を確認、EF CoreマイグレーションをDBへ適用。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set "HOST="
set "PORT="
set "PROJECT="
set /p "HOST=DB ホスト名: "
set /p "PORT=DB ポート番号: "
set /p "PROJECT=EF Core プロジェクトのフォルダ（空欄なら現在のフォルダ）: "
if "%PROJECT%"=="" set "PROJECT=%CD%"

powershell -NoProfile -Command "if($env:HOST -notmatch '^[A-Za-z0-9.-]+$' -or $env:PORT -notmatch '^\d{1,5}$' -or [int]$env:PORT -lt 1 -or [int]$env:PORT -gt 65535){throw 'ホストまたはポートが不正です'}; if(-not (Test-Path -LiteralPath $env:PROJECT)){throw 'プロジェクトフォルダが見つかりません'}"
if errorlevel 1 exit /b 1

powershell -NoProfile -Command "$test=Test-NetConnection -ComputerName $env:HOST -Port ([int]$env:PORT) -InformationLevel Detailed; if(-not $test.TcpTestSucceeded){throw 'DBポートへの接続に失敗しました'}"

if errorlevel 1 (
    echo 接続テストに失敗しました。

    pause
    exit /b 1
)

echo ここでは EF Core の dotnet-ef migration を実行します。
echo 本番DBではなく、対象の接続設定を必ず確認してください。
set "CONFIRM="
set /p "CONFIRM=DBスキーマを変更する場合は MIGRATE と入力: "
if not "%CONFIRM%"=="MIGRATE" exit /b 0
pushd "%PROJECT%"

dotnet ef database update
set "RESULT=%ERRORLEVEL%"
popd
if not "%RESULT%"=="0" echo マイグレーションに失敗しました。DB状態を確認してください。

pause
exit /b %RESULT%
