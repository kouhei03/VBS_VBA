Attribute VB_Name = "Mod88_WBS進捗"
Option Explicit

Public Sub WBS進捗率遅延日数算出()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    ws.Cells(1, 5).Value = "進捗率"
    ws.Cells(1, 6).Value = "遅延日数"
    Dim r As Long
    For r = 2 To lastRow
        If IsNumeric(ws.Cells(r, 3).Value) Then
            ws.Cells(r, 5).Value = CDbl(ws.Cells(r, 3).Value) / 100
            ws.Cells(r, 5).NumberFormat = "0%"
        End If
        If IsDate(ws.Cells(r, 4).Value) Then
            ws.Cells(r, 6).Value = WorksheetFunction.Max(0, DateDiff("d", CDate(ws.Cells(r, 4).Value), Date))
        End If
    Next r
End Sub
