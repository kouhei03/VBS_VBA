Attribute VB_Name = "Mod105_議題集約"
Option Explicit

Public Sub 会議レビュー定例議題アクション集約()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim output As Worksheet: Set output = GetOrCreateSheet
    output.Cells.Clear
    output.Range("A1:D1").Value = Array("会議名", "議題", "アクション", "担当者")
    ws.UsedRange.Copy output.Range("A2")
    output.Columns.AutoFit
End Sub

Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("議題アクション")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "議題アクション"
    End If
End Function
