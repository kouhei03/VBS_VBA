Attribute VB_Name = "Mod100_承認待ち滞留"
Option Explicit

Public Sub 承認待ち滞留案件滞留日数カウント()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    ws.Cells(1, 4).Value = "滞留日数"
    Dim r As Long
    For r = 2 To lastRow
        If CStr(ws.Cells(r, 3).Value) = "承認待ち" And IsDate(ws.Cells(r, 2).Value) Then
            ws.Cells(r, 4).Value = DateDiff("d", CDate(ws.Cells(r, 2).Value), Date)
        End If
    Next r
End Sub
