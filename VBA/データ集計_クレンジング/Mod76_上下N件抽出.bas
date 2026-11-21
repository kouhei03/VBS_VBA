Attribute VB_Name = "Mod76_上下N件抽出"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod76_上下N件抽出
' - ファイル名: Mod76_上下N件抽出.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub 上位下位N件を抽出() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub 上位下位N件を抽出()
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim valueCol As Long: valueCol = CLng(InputBox("並べ替え値の列番号", "上下N件抽出", "3"))
    Dim n As Long: n = CLng(InputBox("抽出件数", "上下N件抽出", "10"))
    Dim out As Worksheet: Set out = GetOrCreateSheet(ws.Parent, "上下N件抽出")
    out.Cells.Clear
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim lastRow As Long: lastRow = GetLastRow(ws)
    Dim lastCol As Long: lastCol = GetLastCol(ws)
    ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Copy out.Range("A1")
    out.UsedRange.Sort Key1:=out.Cells(1, valueCol), Header:=xlYes, Order1:=xlDescending
    out.Cells(1, lastCol + 1).Value = "区分"
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim r As Long
    ' [ブロック解説: 繰り返し処理]
    ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
    For r = 2 To Application.WorksheetFunction.Min(n + 1, lastRow)
        out.Cells(r, lastCol + 1).Value = "上位"
    Next r
    ' [ブロック解説: 繰り返し処理]
    ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
    For r = Application.WorksheetFunction.Max(2, lastRow - n + 1) To lastRow
        out.Cells(r, lastCol + 1).Value = "下位"
    Next r
End Sub

