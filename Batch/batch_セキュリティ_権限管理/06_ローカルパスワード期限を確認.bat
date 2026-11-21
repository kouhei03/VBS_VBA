@echo off
rem 概要: 「ローカルパスワード期限を確認」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "06_ローカルパスワード期限を確認.bat"
rem 実行結果（想定結果）: 「ローカルパスワード期限を確認」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. スクリプト内の設定値を使い、実行条件を確認します。
rem   2. 主処理では、以下のコマンドを使います: Get-LocalUserでローカルアカウントを取得。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal

powershell -NoProfile -Command "$u=Get-LocalUser -Name $env:USERNAME -ErrorAction Stop; Write-Host ('ユーザー: '+$u.Name); Write-Host ('期限切れ: '+$u.PasswordExpired); if($u.PasswordNeverExpires){Write-Warning 'パスワード有効期限は無期限です'}elseif($u.PasswordExpires){Write-Host ('有効期限: '+$u.PasswordExpires); if($u.PasswordExpires -lt (Get-Date).AddDays(14)){Write-Warning 'パスワード期限まで14日未満です'}}else{Write-Host '有効期限を取得できませんでした。ドメインアカウントは別途管理者に確認してください。'}"

pause
