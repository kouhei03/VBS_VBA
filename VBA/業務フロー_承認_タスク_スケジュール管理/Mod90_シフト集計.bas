Attribute VB_Name = "Mod90_シフト集計"
Option Explicit

Public Sub シフト勤務区分集計()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim dic As Object: Set dic = CreateObject("Scripting.Dictionary")
    Dim r As Long
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        Dim typeName As String: typeName = CStr(ws.Cells(r, 3).Value)
        If typeName <> "" Then
            If dic.Exists(typeName) Then dic(typeName) = CLng(dic(typeName)) + 1 Else dic(typeName) = 1
        End If
    Next r
    Dim output As Worksheet: Set output = GetOrCreateSheet
    output.Cells.Clear
    output.Range("A1:B1").Value = Array("勤務区分", "件数")
    Dim outRow As Long: outRow = 2
    Dim key As Variant
    For Each key In dic.Keys
        output.Cells(outRow, 1).Value = key
        output.Cells(outRow, 2).Value = dic(key)
        outRow = outRow + 1
    Next key
End Sub

Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("シフト集計")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "シフト集計"
    End If
End Function
