Attribute VB_Name = "Mod115_完了報告"
Option Explicit

Public Sub 完了報告書ドラフト実績値反映()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim output As Worksheet: Set output = GetOrCreateSheet
    output.Cells.Clear
    output.Cells(1, 1).Value = "完了報告書"
    output.Cells(3, 1).Value = "プロジェクト名"
    output.Cells(3, 2).Value = ws.Cells(2, 1).Value
    output.Cells(4, 1).Value = "実績値"
    output.Cells(4, 2).Value = ws.Cells(2, 3).Value
    output.Cells(5, 1).Value = "作成日"
    output.Cells(5, 2).Value = Date
End Sub

Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("完了報告ドラフト")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "完了報告ドラフト"
    End If
End Function
