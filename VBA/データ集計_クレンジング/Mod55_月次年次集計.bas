Attribute VB_Name = "Mod55_月次年次集計"
Option Explicit

Public Sub 月次集計表を作成()
    Dim dateCol As Long: dateCol = CLng(InputBox("日付列番号", "月次集計", "1"))
    Dim valueCol As Long: valueCol = CLng(InputBox("集計列番号", "月次集計", "3"))
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim out As Worksheet: Set out = GetOrCreateSheet(ws.Parent, "月次集計")
    out.Cells.Clear
    Dim dic As Object: Set dic = CreateObject("Scripting.Dictionary")
    Dim r As Long
    For r = 2 To GetLastRow(ws)
        If IsDate(ws.Cells(r, dateCol).Value) Then
            Dim key As String: key = Format$(CDate(ws.Cells(r, dateCol).Value), "yyyy/mm")
            If Not dic.Exists(key) Then dic.Add key, 0
            If IsNumeric(ws.Cells(r, valueCol).Value) Then dic(key) = dic(key) + CDbl(ws.Cells(r, valueCol).Value)
        End If
    Next r
    out.Cells(1, 1).Value = "年月": out.Cells(1, 2).Value = "合計"
    r = 2
    Dim k As Variant
    For Each k In dic.Keys
        out.Cells(r, 1).Value = k
        out.Cells(r, 2).Value = dic(k)
        r = r + 1
    Next k
End Sub
