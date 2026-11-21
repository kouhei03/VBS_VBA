Attribute VB_Name = "Mod163_期間レポート"
Option Explicit

Public Sub 週次月次年次レポート作成()
    Dim period As String
    period = InputBox("期間を入力: 週次 / 月次 / 年次", "レポート作成")
    If period <> "週次" And period <> "月次" And period <> "年次" Then Exit Sub
    Dim source As Worksheet: Set source = ActiveSheet
    Dim output As Worksheet: Set output = GetOrCreateSheet(period & "レポート")
    output.Cells.Clear
    output.Cells(1, 1).Value = period & "レポート"
    output.Cells(2, 1).Value = "作成日"
    output.Cells(2, 2).Value = Date
    source.UsedRange.Copy output.Range("A4")
    output.Columns.AutoFit
End Sub

Private Function GetOrCreateSheet(ByVal name As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = name
    End If
End Function
