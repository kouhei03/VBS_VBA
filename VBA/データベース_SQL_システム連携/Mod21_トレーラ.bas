Attribute VB_Name = "Mod21_トレーラ"
Option Explicit

Public Sub レコード件数合計金額トレーラ付与()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    ws.Cells(lastRow + 1, 1).Value = "TRAILER"
    ws.Cells(lastRow + 1, 2).Value = lastRow - 1
    ws.Cells(lastRow + 1, 3).Value = Application.WorksheetFunction.Sum(ws.Columns(3))
End Sub
