Attribute VB_Name = "Mod115_テンプレート複製"
Option Explicit

Public Sub テンプレートファイル複製作業ブック生成()
    Dim templatePath As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .Filters.Add "Excel", "*.xls;*.xlsx;*.xlsm"
        If .Show <> -1 Then Exit Sub
        templatePath = .SelectedItems(1)
    End With
    Dim destinationFolder As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show <> -1 Then Exit Sub
        destinationFolder = .SelectedItems(1)
    End With
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim name As String
    name = InputBox("作業ブック名を入力", "テンプレート複製", "作業_" & Format$(Now, "yyyymmdd"))
    If Trim$(name) = "" Then Exit Sub
    Dim extension As String: extension = "." & fso.GetExtensionName(templatePath)
    If LCase$(Right$(name, Len(extension))) <> LCase$(extension) Then name = name & extension
    fso.CopyFile templatePath, fso.BuildPath(destinationFolder, name), True
    Workbooks.Open fso.BuildPath(destinationFolder, name)
End Sub
