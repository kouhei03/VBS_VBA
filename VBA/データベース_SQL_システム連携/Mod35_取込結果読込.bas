Attribute VB_Name = "Mod35_取込結果読込"
Option Explicit

Public Sub エラー取込結果ファイル読込集計()
    Dim filePath As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .Filters.Add "CSV", "*.csv"
        If .Show <> -1 Then Exit Sub
        filePath = .SelectedItems(1)
    End With
    Dim temp As Workbook: Set temp = Workbooks.Open(filePath)
    Dim source As Worksheet: Set source = temp.Worksheets(1)
    Dim output As Worksheet: Set output = GetOrCreateSheet
    output.Cells.Clear
    source.UsedRange.Copy output.Range("A1")
    temp.Close False
End Sub

Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("取込結果")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "取込結果"
    End If
End Function
