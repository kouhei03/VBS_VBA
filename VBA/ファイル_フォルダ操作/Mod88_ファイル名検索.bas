Attribute VB_Name = "Mod88_ファイル名検索"
Option Explicit

Public Sub ファイル名条件一致検索()
    Dim folderPath As String, pattern As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show <> -1 Then Exit Sub
        folderPath = .SelectedItems(1)
    End With
    pattern = InputBox("ワイルドカード検索文字を入力（例: 報告*.xlsx）", "ファイル名検索")
    If Trim$(pattern) = "" Then Exit Sub

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim ws As Worksheet: Set ws = GetOrCreateSheet("ファイル名検索")
    ws.Cells.Clear
    Dim r As Long: r = 1
    SearchName fso.GetFolder(folderPath), pattern, ws, r
    ws.Columns("A:B").AutoFit
End Sub

Private Sub SearchName(ByVal folder As Object, ByVal pattern As String, ByVal ws As Worksheet, ByRef r As Long)
    Dim file As Object
    For Each file In folder.Files
        If file.Name Like pattern Then
            ws.Cells(r, 1).Value = file.Name
            ws.Cells(r, 2).Value = file.Path
            r = r + 1
        End If
    Next file
    Dim subfolder As Object
    For Each subfolder In folder.SubFolders
        SearchName subfolder, pattern, ws, r
    Next subfolder
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
