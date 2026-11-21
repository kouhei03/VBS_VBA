Attribute VB_Name = "Mod76_上下N件抽出"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 上位下位N件を抽出を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「上位下位N件を抽出」を選択して実行する。

' ── 上位下位N件を抽出 ──
Public Sub 上位下位N件を抽出()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim valueCol As Long: valueCol = CLng(InputBox("並べ替え値の列番号", "上下N件抽出", "3"))
    Dim n As Long: n = CLng(InputBox("抽出件数", "上下N件抽出", "10"))
    Dim out As Worksheet: Set out = GetOrCreateSheet(ws.Parent, "上下N件抽出")
    out.Cells.Clear
    Dim lastRow As Long: lastRow = GetLastRow(ws)
    Dim lastCol As Long: lastCol = GetLastCol(ws)
    ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Copy out.Range("A1")
    out.UsedRange.Sort Key1:=out.Cells(1, valueCol), Header:=xlYes, Order1:=xlDescending
    out.Cells(1, lastCol + 1).Value = "区分"
    Dim r As Long
    For r = 2 To Application.WorksheetFunction.Min(n + 1, lastRow)
        out.Cells(r, lastCol + 1).Value = "上位"
    Next r
    For r = Application.WorksheetFunction.Max(2, lastRow - n + 1) To lastRow
        out.Cells(r, lastCol + 1).Value = "下位"
    Next r
End Sub




