Attribute VB_Name = "Mod44_空白列削除"
Option Explicit

Public Sub 空白列を一括削除()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim lastRow As Long: lastRow = GetLastRow(ws)
    Dim lastCol As Long: lastCol = GetLastCol(ws)
    Dim delRange As Range
    Dim c As Long
    For c = lastCol To 1 Step -1
        If Application.WorksheetFunction.CountA(ws.Range(ws.Cells(1, c), ws.Cells(lastRow, c))) = 0 Then
            If delRange Is Nothing Then Set delRange = ws.Columns(c) Else Set delRange = Union(delRange, ws.Columns(c))
        End If
    Next c
    If Not delRange Is Nothing Then delRange.Delete
End Sub
