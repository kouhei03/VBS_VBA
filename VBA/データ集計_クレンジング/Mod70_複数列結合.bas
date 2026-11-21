Attribute VB_Name = "Mod70_複数列結合"
Option Explicit

Public Sub 複数列を結合()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim colText As String: colText = InputBox("結合する列番号（例: 1,2,3）", "列結合")
    If colText = "" Then Exit Sub
    Dim separator As String: separator = InputBox("区切り文字", "列結合", "")
    Dim parts As Variant: parts = Split(colText, ",")
    Dim r As Long, i As Long
    Dim targetCol As Long: targetCol = GetLastCol(ws) + 1
    ws.Cells(1, targetCol).Value = "結合結果"
    For r = 2 To GetLastRow(ws)
        Dim result As String
        For i = LBound(parts) To UBound(parts)
            If Len(result) > 0 Then result = result & separator
            result = result & CStr(ws.Cells(r, CLng(Trim$(CStr(parts(i))))).Value)
        Next i
        ws.Cells(r, targetCol).Value = result
    Next r
End Sub
