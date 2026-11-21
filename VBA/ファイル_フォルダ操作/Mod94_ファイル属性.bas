Attribute VB_Name = "Mod94_ファイル属性"
Option Explicit

Public Sub ファイル更新作成日時サイズ取得()
    Dim filePath As String
    With Application.FileDialog(msoFileDialogFilePicker)
        If .Show <> -1 Then Exit Sub
        filePath = .SelectedItems(1)
    End With
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim file As Object: Set file = fso.GetFile(filePath)
    Dim ws As Worksheet: Set ws = GetOrCreateSheet("ファイル属性")
    ws.Cells.Clear
    ws.Cells(1, 1).Value = "項目"
    ws.Cells(1, 2).Value = "値"
    ws.Cells(2, 1).Value = "名前"
    ws.Cells(2, 2).Value = file.Name
    ws.Cells(3, 1).Value = "作成日時"
    ws.Cells(3, 2).Value = file.DateCreated
    ws.Cells(4, 1).Value = "更新日時"
    ws.Cells(4, 2).Value = file.DateLastModified
    ws.Cells(5, 1).Value = "サイズ"
    ws.Cells(5, 2).Value = file.Size
    ws.Cells(6, 1).Value = "パス"
    ws.Cells(6, 2).Value = file.Path
    ws.Columns("A:B").AutoFit
End Sub

Private Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = sheetName
    End If
End Function
