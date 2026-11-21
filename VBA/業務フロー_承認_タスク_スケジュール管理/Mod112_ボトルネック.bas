Attribute VB_Name = "Mod112_ボトルネック"
Option Explicit

Public Sub 業務フロー所要時間ボトルネック分析()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim output As Worksheet: Set output = GetOrCreateSheet
    output.Cells.Clear
    output.Range("A1:C1").Value = Array("工程", "平均所要時間", "件数")
    Dim dic As Object: Set dic = CreateObject("Scripting.Dictionary")
    Dim r As Long
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        Dim stepName As String: stepName = CStr(ws.Cells(r, 1).Value)
        If IsNumeric(ws.Cells(r, 2).Value) Then
            If dic.Exists(stepName) Then dic(stepName) = CDbl(dic(stepName)) + CDbl(ws.Cells(r, 2).Value) _
            Else dic(stepName) = CDbl(ws.Cells(r, 2).Value)
        End If
    Next r
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
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("ボトルネック")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "ボトルネック"
    End If
End Function
