Attribute VB_Name = "Mod43_空白行削除"
Option Explicit

Public Sub 空白行を一括削除()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim lastRow As Long: lastRow = GetLastRow(ws)
    Dim lastCol As Long: lastCol = GetLastCol(ws)
    Dim delRange As Range
    Dim r As Long
    For r = lastRow To 2 Step -1
        If Application.WorksheetFunction.CountA(ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol))) = 0 Then
            If delRange Is Nothing Then Set delRange = ws.Rows(r) Else Set delRange = Union(delRange, ws.Rows(r))
        End If
    Next r
    If Not delRange Is Nothing Then delRange.Delete
End Sub
