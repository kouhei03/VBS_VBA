Attribute VB_Name = "Mod151_システム情報"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod151_システム情報
' - ファイル名: Mod151_システム情報.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub Windowsシステム情報環境変数取得() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub Windowsシステム情報環境変数取得()
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim shell As Object: Set shell = CreateObject("WScript.Shell")
    Dim environment As Object: Set environment = shell.Environment("PROCESS")
    Dim ws As Worksheet: Set ws = GetOrCreateSheet(ThisWorkbook, "システム情報")
    ws.Cells.Clear
    ws.Cells(1, 1).Value = "項目"
    ws.Cells(1, 2).Value = "値"
    ws.Cells(2, 1).Value = "COMPUTERNAME"
    ws.Cells(2, 2).Value = environment("COMPUTERNAME")
    ws.Cells(3, 1).Value = "USERNAME"
    ws.Cells(3, 2).Value = environment("USERNAME")
    ws.Cells(4, 1).Value = "SystemRoot"
    ws.Cells(4, 2).Value = environment("SystemRoot")
    ws.Cells(5, 1).Value = "TEMP"
    ws.Cells(5, 2).Value = environment("TEMP")
    ws.Columns("A:B").AutoFit
End Sub

