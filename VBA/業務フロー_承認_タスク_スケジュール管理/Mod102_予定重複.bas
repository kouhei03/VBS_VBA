Attribute VB_Name = "Mod102_予定重複"
Option Explicit

Public Sub スケジュール予定重複検知()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim r As Long
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        Dim compareRow As Long
        For compareRow = r + 1 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            If CStr(ws.Cells(r, 1).Value) = CStr(ws.Cells(compareRow, 1).Value) _
               And CDate(ws.Cells(r, 2).Value) < CDate(ws.Cells(compareRow, 3).Value) _
               And CDate(ws.Cells(compareRow, 2).Value) < CDate(ws.Cells(r, 3).Value) Then
                ws.Rows(r).Interior.Color = RGB(255, 235, 156)
                ws.Rows(compareRow).Interior.Color = RGB(255, 235, 156)
            End If
        Next compareRow
    Next r
End Sub
