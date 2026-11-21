Attribute VB_Name = "Mod119_ワークフローアーカイブ"
Option Explicit

Public Sub ワークフロー履歴アーカイブ世代管理()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("ワークフロー履歴")
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "ワークフロー履歴シートがありません。", vbExclamation
        Exit Sub
    End If
    Dim archive As Worksheet: Set archive = GetOrCreateSheet
    ws.UsedRange.Copy archive.Range("A1")
    ws.Cells.Clear
End Sub

Private Function GetOrCreateSheet() As Worksheet
    Dim name As String: name = "履歴_" & Format$(Now, "yyyymmdd_hhnnss")
    Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
    GetOrCreateSheet.Name = name
End Function
