Attribute VB_Name = "Mod87_拡張子抽出"
Option Explicit

Public Sub 特定拡張子ファイル抽出()
    Dim folderPath As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show <> -1 Then Exit Sub
        folderPath = .SelectedItems(1)
    End With
    Dim extensions As String
    extensions = InputBox("抽出する拡張子をカンマ区切りで入力（例: xlsx,xlsm,csv,pdf）", "拡張子抽出")
    If Trim$(extensions) = "" Then Exit Sub

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim allowed As Object: Set allowed = CreateObject("Scripting.Dictionary")
    allowed.CompareMode = vbTextCompare
    Dim item As Variant
    For Each item In Split(extensions, ",")
        allowed.Add Replace(Trim$(CStr(item)), ".", ""), True
    Next item

    Dim ws As Worksheet: Set ws = GetOrCreateSheet("拡張子抽出")
    ws.Cells.Clear
    Dim r As Long: r = 1
    ExtractFiles fso.GetFolder(folderPath), allowed, ws, r
    ws.Columns("A:B").AutoFit
End Sub

Private Sub ExtractFiles(ByVal folder As Object, ByVal allowed As Object, ByVal ws As Worksheet, ByRef r As Long)
    Dim file As Object
    For Each file In folder.Files
        If allowed.Exists(LCase$(Mid$(file.Name, InStrRev(file.Name, ".") + 1))) Then
            ws.Cells(r, 1).Value = file.Name
            ws.Cells(r, 2).Value = file.Path
            r = r + 1
        End If
    Next file
    Dim subfolder As Object
    For Each subfolder In folder.SubFolders
        ExtractFiles subfolder, allowed, ws, r
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
