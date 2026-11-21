@echo off
rem 概要: 「単一ソースをコンパイル」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "07_単一ソースをコンパイル.bat"
rem 実行結果（想定結果）: 「単一ソースをコンパイル」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では本文内のWindowsまたはPowerShellコマンドを実行します。分岐がある場合は選択条件に従います。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
echo 1: C++（g++）
echo 2: Java（javac）
echo 3: Rust（rustc）
set "LANGUAGE="
set /p "LANGUAGE=言語を選択してください (1-3): "
set "SOURCE="
set /p "SOURCE=コンパイルするソースファイルのフルパス: "

if not exist "%SOURCE%" (
    echo ソースファイルが見つかりません。

    pause
    exit /b 1
)

for %%F in ("%SOURCE%") do set "OUTDIR=%%~dpF"
if "%LANGUAGE%"=="1" goto compile_cpp
if "%LANGUAGE%"=="2" goto compile_java
if "%LANGUAGE%"=="3" goto compile_rust
echo 言語が選択されていません。
exit /b 1

:compile_cpp

for %%F in ("%SOURCE%") do set "OUTPUT=%%~dpnF.exe"
g++ -std=c++17 "%SOURCE%" -o "%OUTPUT%"
goto result

:compile_java

if not exist "%OUTDIR%out" mkdir "%OUTDIR%out"
javac -d "%OUTDIR%out" "%SOURCE%"
goto result

:compile_rust

for %%F in ("%SOURCE%") do set "OUTPUT=%%~dpnF.exe"
rustc "%SOURCE%" -o "%OUTPUT%"

:result

if errorlevel 1 echo コンパイルに失敗しました。

pause
