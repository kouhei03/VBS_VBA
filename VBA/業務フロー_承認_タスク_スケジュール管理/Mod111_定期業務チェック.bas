Attribute VB_Name = "Mod111_定期業務チェック"
Option Explicit

Public Sub 定期業務日次週次月次年次チェックリスト生成()
    Dim cycle As String
    cycle = InputBox("周期を入力: 日次 / 週次 / 月次 / 年次", "定期業務")
    If cycle = "" Then Exit Sub
    Dim ws As Worksheet: Set ws = GetOrCreateSheet
    Dim r As Long: r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(r, 1).Value = cycle
    ws.Cells(r, 2).Value = Format$(Now, "yyyy/mm/dd hh:mm")
    ws.Cells(r, 3).Value = "未実施"
End Sub

Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("定期業務チェック")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "定期業務チェック"
    End If
End Function
