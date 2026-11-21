Attribute VB_Name = "Mod97_承認ルート"
Option Explicit

Public Sub 金額権限規定承認ルート承認者判定()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim r As Long
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If IsNumeric(ws.Cells(r, 2).Value) Then
            If CDbl(ws.Cells(r, 2).Value) >= 1000000 Then
                ws.Cells(r, 3).Value = "役員承認"
            ElseIf CDbl(ws.Cells(r, 2).Value) >= 100000 Then
                ws.Cells(r, 3).Value = "部門長承認"
            Else
                ws.Cells(r, 3).Value = "課長承認"
            End If
        End If
    Next r
End Sub
