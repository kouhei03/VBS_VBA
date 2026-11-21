Attribute VB_Name = "Mod110_SLA違反"
Option Explicit

Public Sub SLA違反ニアミス抽出()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim r As Long
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If IsDate(ws.Cells(r, 3).Value) Then
            Dim remaining As Double: remaining = CDate(ws.Cells(r, 3).Value) - Now
            If remaining < 0 Then
                ws.Cells(r, 4).Value = "SLA違反"
                ws.Rows(r).Interior.Color = RGB(255, 199, 206)
            ElseIf remaining <= 1 Then
                ws.Cells(r, 4).Value = "ニアミス"
                ws.Rows(r).Interior.Color = RGB(255, 235, 156)
            End If
        End If
    Next r
End Sub
