@echo off
rem 概要: 「無操作時間後に画面ロック」を行うBATです。
rem 実行方法: 対象フォルダでダブルクリックするか、コマンドプロンプトから実行。
rem 実行コマンド例: call "10_無操作時間後に画面ロック.bat"
rem 実行結果（想定結果）: 「無操作時間後に画面ロック」の処理結果が画面表示、ファイル出力、または対象の操作として反映されます。入力条件や外部ツールの状態により異なる場合があります。
rem 実行環境: Windows 10/11、コマンドプロンプト。必要な権限・PowerShell・外部ツールは本文の案内を参照。
rem コード詳細解説:
rem   1. 対話入力（set /p／choice）で対象や条件を受け取ります。入力値・対象の有無・確認応答に応じて分岐します。
rem   2. 主処理では、以下のコマンドを使います: PowerShellのAdd-Typeで対象を処理、PowerShellのNew-Objectで対象を処理。複数の分岐がある場合は条件に応じて実行します。
rem   3. 実行結果は画面表示または指定先への出力・対象の変更として反映されます。終了コードや異常条件を確認し、失敗時は通知または中断します。

chcp 65001 >nul
setlocal
set "MINUTES="
set /p "MINUTES=無操作の時間（分）: "
echo 監視中です。停止は Ctrl+C です。

powershell -NoProfile -Command "$minutes=0; if(-not [int]::TryParse($env:MINUTES,[ref]$minutes) -or $minutes -lt 1){throw '1分以上で指定してください'}; $q=[char]34; $source='using System; using System.Runtime.InteropServices; public static class IdleLockApi { [StructLayout(LayoutKind.Sequential)] public struct LASTINPUTINFO { public uint cbSize; public uint dwTime; } [DllImport('+$q+'user32.dll'+$q+')] public static extern bool GetLastInputInfo(ref LASTINPUTINFO p); [DllImport('+$q+'kernel32.dll'+$q+')] public static extern uint GetTickCount(); [DllImport('+$q+'user32.dll'+$q+')] public static extern bool LockWorkStation(); }'; Add-Type -TypeDefinition $source; $limit=$minutes*60000; while($true){$info=New-Object 'IdleLockApi+LASTINPUTINFO'; $info.cbSize=[Runtime.InteropServices.Marshal]::SizeOf($info); [void][IdleLockApi]::GetLastInputInfo([ref]$info); $idle=[uint32]([IdleLockApi]::GetTickCount()-$info.dwTime); if($idle -ge $limit){Write-Host '無操作時間を超えたため画面をロックします'; [void][IdleLockApi]::LockWorkStation(); Start-Sleep -Seconds 60}; Start-Sleep -Seconds 5}"
