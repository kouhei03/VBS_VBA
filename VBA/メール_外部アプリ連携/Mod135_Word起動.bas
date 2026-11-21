Attribute VB_Name = "Mod135_Word起動"
Option Explicit

Public Sub Word文書起動新規作成保存()
    Dim word As Object
    On Error Resume Next
    Set word = GetObject(, "Word.Application")
    On Error GoTo 0
    If word Is Nothing Then Set word = CreateObject("Word.Application")
    word.Visible = True
    Dim document As Object: Set document = word.Documents.Add
    Dim filePath As String
    With Application.FileDialog(msoFileDialogSaveAs)
        .Filters.Clear
        .Filters.Add "Word", "*.docx"
        If .Show = -1 Then filePath = .SelectedItems(1)
    End With
    If filePath <> "" Then document.SaveAs2 filePath
End Sub
