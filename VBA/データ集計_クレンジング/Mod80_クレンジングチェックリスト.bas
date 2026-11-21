Attribute VB_Name = "Mod80_クレンジングチェックリスト"
Option Explicit

Public Sub クレンジング前後チェックリスト作成()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim out As Worksheet: Set out = GetOrCreateSheet(ws.Parent, "クレンジングチェック")
    out.Cells.Clear
    out.Cells(1, 1).Value = "項目"
    out.Cells(1, 2).Value = "件数"
    out.Cells(1, 3).Value = "合計"
    out.Cells(2, 1).Value = "全行"
    out.Cells(2, 2).Value = GetLastRow(ws) - 1
    out.Cells(3, 1).Value = "空白セル"
    out.Cells(3, 2).Value = Application.CountBlank(ws.UsedRange)
    out.Cells(4, 1).Value = "数値合計"
    out.Cells(4, 3).Value = Application.WorksheetFunction.Sum(ws.UsedRange)
End Sub
