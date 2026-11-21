Attribute VB_Name = "Mod168_小計合計書式"
Option Explicit

Public Sub 合計小計行の太線二重線塗りつぶし()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim r As Long
    For r = 1 To lastRow
        If InStr(CStr(ws.Cells(r, 1).Value), "合計") > 0 Or _
           InStr(CStr(ws.Cells(r, 1).Value), "小計") > 0 Then
            ws.Rows(r).Font.Bold = True
            ws.Rows(r).Borders(xlTop).LineStyle = xlDouble
            ws.Rows(r).Interior.Color = RGB(217, 234, 247)
        End If
    Next r
End Sub
