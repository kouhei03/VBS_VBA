Attribute VB_Name = "Mod167_ゼブラ"
Option Explicit

Public Sub 交互行ゼブラカラー適用()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim lastColumn As Long: lastColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Dim r As Long
    For r = 2 To lastRow
        If r Mod 2 = 0 Then ws.Range(ws.Cells(r, 1), ws.Cells(r, lastColumn)).Interior.Color = RGB(242, 242, 242)
    Next r
End Sub
