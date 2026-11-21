Attribute VB_Name = "Mod113_空フォルダ"
Option Explicit

Public Sub 空フォルダ検出削除()
    Dim folderPath As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show <> -1 Then Exit Sub
        folderPath = .SelectedItems(1)
    End With
    Dim deleteEmpty As Boolean
    deleteEmpty = MsgBox("空フォルダを検出し削除しますか？", vbYesNo + vbQuestion) = vbYes
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim ws As Worksheet: Set ws = GetOrCreateSheet("空フォルダ")
    ws.Cells.Clear
    Dim r As Long: r = 1
    FindEmptyFolders fso.GetFolder(folderPath), ws, r, deleteEmpty
    ws.Columns("A:A").AutoFit
End Sub

Private Sub FindEmptyFolders(ByVal folder As Object, ByVal ws As Worksheet, ByRef r As Long, ByVal deleteEmpty As Boolean)
    Dim subfolder As Object
    For Each subfolder In folder.SubFolders
        FindEmptyFolders subfolder, ws, r, deleteEmpty
    Next subfolder
    If folder.Files.Count = 0 And folder.SubFolders.Count = 0 Then
        ws.Cells(r, 1).Value = folder.Path
        r = r + 1
        If deleteEmpty Then folder.Delete True
    End If
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
