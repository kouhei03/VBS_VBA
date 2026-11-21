Attribute VB_Name = "Mod64_クロス集計"
Option Explicit

Public Sub クロス集計表を作成()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim rowCol As Long: rowCol = CLng(InputBox("行項目の列番号", "クロス集計", "1"))
    Dim colCol As Long: colCol = CLng(InputBox("列項目の列番号", "クロス集計", "2"))
    Dim valueCol As Long: valueCol = CLng(InputBox("集計値の列番号", "クロス集計", "3"))
    Dim out As Worksheet: Set out = GetOrCreateSheet(ws.Parent, "クロス集計")
    out.Cells.Clear

    Dim rowsDic As Object: Set rowsDic = CreateObject("Scripting.Dictionary")
    Dim colsDic As Object: Set colsDic = CreateObject("Scripting.Dictionary")
    Dim sums As Object: Set sums = CreateObject("Scripting.Dictionary")
    Dim r As Long
    For r = 2 To GetLastRow(ws)
        Dim rowKey As String: rowKey = CStr(ws.Cells(r, rowCol).Value)
        Dim colKey As String: colKey = CStr(ws.Cells(r, colCol).Value)
        rowsDic(rowKey) = True: colsDic(colKey) = True
        Dim key As String: key = rowKey & vbNullChar & colKey
        If Not sums.Exists(key) Then sums.Add key, 0
        If IsNumeric(ws.Cells(r, valueCol).Value) Then sums(key) = sums(key) + CDbl(ws.Cells(r, valueCol).Value)
    Next r
    Dim c As Long: c = 2
    Dim colName As Variant
    For Each colName In colsDic.Keys
        out.Cells(1, c).Value = colName
        c = c + 1
    Next colName
    r = 2
    Dim rowName As Variant
    For Each rowName In rowsDic.Keys
        out.Cells(r, 1).Value = rowName
        c = 2
        For Each colName In colsDic.Keys
            key = CStr(rowName) & vbNullChar & CStr(colName)
            out.Cells(r, c).Value = IIf(sums.Exists(key), sums(key), 0)
            c = c + 1
        Next colName
        r = r + 1
    Next rowName
End Sub
