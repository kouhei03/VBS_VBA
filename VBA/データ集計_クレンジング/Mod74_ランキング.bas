Attribute VB_Name = "Mod74_ランキング"
Option Explicit

Public Sub ランキング順位を付与()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim valueCol As Long: valueCol = CLng(InputBox("ランキング対象列番号", "ランキング", "3"))
    Dim rankCol As Long: rankCol = GetLastCol(ws) + 1
    ws.Cells(1, rankCol).Value = "ランキング"
    Dim r As Long
    For r = 2 To GetLastRow(ws)
        If IsNumeric(ws.Cells(r, valueCol).Value) Then
            ws.Cells(r, rankCol).Formula = "=RANK(" & ws.Cells(r, valueCol).Address(False, False) & "," & _
                                           ws.Range(ws.Cells(2, valueCol), ws.Cells(GetLastRow(ws), valueCol)).Address(True, True) & ",0)"
        End If
    Next r
End Sub
