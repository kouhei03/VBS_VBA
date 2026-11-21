Attribute VB_Name = "Mod86_定型タスク展開"
Option Explicit

Public Sub 週次月次タスクテンプレート展開()
    Dim cycle As String
    cycle = InputBox("周期を入力: 週次 / 月次", "定型タスク")
    If cycle <> "週次" And cycle <> "月次" Then Exit Sub
    Dim ws As Worksheet: Set ws = GetOrCreateSheet
    Dim r As Long: r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(r, 1).Value = Format$(Now, "yyyy/mm/dd")
    ws.Cells(r, 2).Value = cycle & "定期業務"
    ws.Cells(r, 3).Value = "未着手"
End Sub

Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("定型タスク")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "定型タスク"
    End If
End Function
