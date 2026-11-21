Attribute VB_Name = "Mod84_担当者負荷"
Option Explicit

Public Sub 担当者別未完了タスク負荷集計()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim dic As Object: Set dic = CreateObject("Scripting.Dictionary")
    Dim r As Long
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        Dim owner As String: owner = CStr(ws.Cells(r, 2).Value)
        Dim status As String: status = CStr(ws.Cells(r, 3).Value)
        If owner <> "" And status <> "完了" Then
            If dic.Exists(owner) Then dic(owner) = CLng(dic(owner)) + 1 Else dic(owner) = 1
        End If
    Next r
    Dim output As Worksheet: Set output = GetOrCreateSheet
    output.Cells.Clear
    output.Range("A1:B1").Value = Array("担当者", "未完了件数")
    Dim outRow As Long: outRow = 2
    Dim key As Variant
    For Each key In dic.Keys
        output.Cells(outRow, 1).Value = key
        output.Cells(outRow, 2).Value = dic(key)
        outRow = outRow + 1
    Next key
    output.Columns("A:B").AutoFit
End Sub

Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("担当者負荷")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "担当者負荷"
    End If
End Function
