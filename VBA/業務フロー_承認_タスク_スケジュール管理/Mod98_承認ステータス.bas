Attribute VB_Name = "Mod98_承認ステータス"
Option Explicit

Public Sub 承認却下差戻ステータス履歴更新()
    Dim ws As Worksheet: Set ws = ActiveSheet
    If Selection Is Nothing Then Exit Sub
    Dim status As String
    status = InputBox("ステータスを入力（承認 / 却下 / 差し戻し）", "承認ステータス")
    Selection.Value = status
    Dim log As Worksheet: Set log = GetOrCreateSheet
    Dim r As Long: r = log.Cells(log.Rows.Count, 1).End(xlUp).Row + 1
    log.Cells(r, 1).Value = Now
    log.Cells(r, 2).Value = Selection.Address
    log.Cells(r, 3).Value = status
End Sub

Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("承認履歴")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "承認履歴"
    End If
End Function
