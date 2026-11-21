Attribute VB_Name = "Mod108_引継ぎ"
Option Explicit

Public Sub 業務引継ぎ引継チェックリスト進捗管理()
    Dim ws As Worksheet: Set ws = GetOrCreateSheet
    ws.Cells.Clear
    ws.Range("A1:D1").Value = Array("引継項目", "担当者", "期限", "完了")
    Dim r As Long: r = 2
    For r = 2 To ActiveSheet.Cells(ActiveSheet.Rows.Count, 1).End(xlUp).Row
        ActiveSheet.Rows(r).Copy ws.Rows(r)
    Next r
End Sub

Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("引継チェックリスト")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "引継チェックリスト"
    End If
End Function
