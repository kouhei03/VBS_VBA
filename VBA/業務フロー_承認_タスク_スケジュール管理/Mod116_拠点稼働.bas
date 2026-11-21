Attribute VB_Name = "Mod116_拠点稼働"
Option Explicit

Public Sub 複数拠点チーム稼働状況一覧()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim dic As Object: Set dic = CreateObject("Scripting.Dictionary")
    Dim r As Long
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        Dim key As String: key = CStr(ws.Cells(r, 1).Value) & "/" & CStr(ws.Cells(r, 2).Value)
        If dic.Exists(key) Then dic(key) = CLng(dic(key)) + 1 Else dic(key) = 1
    Next r
    Dim output As Worksheet: Set output = GetOrCreateSheet
    output.Cells.Clear
    Dim outRow As Long: outRow = 1
    Dim item As Variant
    For Each item In dic.Keys
        output.Cells(outRow, 1).Value = item
        output.Cells(outRow, 2).Value = dic(item)
        outRow = outRow + 1
    Next item
End Sub

Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("拠点稼働")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "拠点稼働"
    End If
End Function
