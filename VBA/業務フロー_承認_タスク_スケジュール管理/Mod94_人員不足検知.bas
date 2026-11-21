Attribute VB_Name = "Mod94_人員不足検知"
Option Explicit

Public Sub 時間帯日別人員不足過不足検知()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim r As Long
    For r = 2 To lastRow
        If IsNumeric(ws.Cells(r, 3).Value) And IsNumeric(ws.Cells(r, 4).Value) Then
            If CDbl(ws.Cells(r, 3).Value) < CDbl(ws.Cells(r, 4).Value) Then
                ws.Cells(r, 5).Value = "不足"
                ws.Rows(r).Interior.Color = RGB(255, 199, 206)
            Else
                ws.Cells(r, 5).Value = "十分"
            End If
        End If
    Next r
End Sub
