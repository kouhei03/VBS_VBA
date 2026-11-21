@echo off
rem 概要: 「CSV列を抽出」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "03_CSV列を抽出.bat"
rem 実行結果（想定結果）: 「CSV列を抽出」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: PowerShellのImport-Csvで対象を処理、Export-CsvでCSVファイルを出力。複数の分岐がある場合は条件に応じて実行します。
rem   3. 結果やログを本文で指定したファイルへ保存し、必要に応じて画面にも表示します。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set "CSV="
set "COLUMN="
set "OUTPUT="
set /p "CSV=入力 CSV ファイル: "
set /p "COLUMN=抽出する列名（ヘッダー）: "
set /p "OUTPUT=出力 CSV ファイル: "

if not exist "%CSV%" (
    echo CSV ファイルが見つかりません。

    pause
    exit /b 1
)

if /i "%CSV%"=="%OUTPUT%" (
    echo 入力 CSV と出力 CSV は別にしてください。
    exit /b 1
)

powershell -NoProfile -Command "$ErrorActionPreference='Stop'; $rows=@(Import-Csv -LiteralPath $env:CSV); if($rows.Count -and $rows[0].PSObject.Properties.Name -notcontains $env:COLUMN){throw '指定列が見つかりません'}; $rows | Select-Object -Property $env:COLUMN | Export-Csv -LiteralPath $env:OUTPUT -NoTypeInformation -Encoding UTF8"
if errorlevel 1 echo CSV の抽出に失敗しました。

pause
