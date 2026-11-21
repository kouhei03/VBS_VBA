Attribute VB_Name = "Mod90_ファイル移動"
Option Explicit

Public Sub ファイルを移動()
    Dim sourcePath As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Sub
        sourcePath = .SelectedItems(1)
    End With
    Dim destinationFolder As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "移動先フォルダを選択"
        If .Show <> -1 Then Exit Sub
        destinationFolder = .SelectedItems(1)
    End With
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(destinationFolder) Then fso.CreateFolder destinationFolder
    Dim destinationPath As String
    destinationPath = fso.BuildPath(destinationFolder, fso.GetFileName(sourcePath))
    If fso.FileExists(destinationPath) Then
        If MsgBox("同名ファイルがあります。上書きしますか？", vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    End If
    fso.MoveFile sourcePath, destinationPath
    MsgBox "移動しました。", vbInformation
End Sub
