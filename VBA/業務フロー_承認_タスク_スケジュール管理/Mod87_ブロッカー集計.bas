Attribute VB_Name = "Mod87_ブロッカー集計"
Option Explicit

Public Sub 遅延理由ブロッカー入力促し集計()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim output As Worksheet: Set output = GetOrCreateSheet
    output.Cells.Clear
    output.Range("A1:C1").Value = Array("案件", "遅延理由", "件数")
    Dim dic As Object: Set dic = CreateObject("Scripting.Dictionary")
    Dim r As Long
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        Dim reason As String: reason = CStr(ws.Cells(r, 4).Value)
        If reason = "" Then reason = "未入力"
        If dic.Exists(reason) Then dic(reason) = CLng(dic(reason)) + 1 Else dic(reason) = 1
    Next r
    Dim outRow As Long: outRow = 2
    Dim key As Variant
    For Each key In dic.Keys
        output.Cells(outRow, 2).Value = key
        output.Cells(outRow, 3).Value = dic(key)
        outRow = outRow + 1
    Next key
End Sub

Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("ブロッカー集計")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "ブロッカー集計"
    End If
End Function
