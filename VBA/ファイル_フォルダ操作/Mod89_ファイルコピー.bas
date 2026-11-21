Attribute VB_Name = "Mod89_ファイルコピー"
Option Explicit

Public Sub ファイルをコピー()
    Dim sourcePath As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Sub
        sourcePath = .SelectedItems(1)
    End With
    Dim destinationFolder As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "コピー先フォルダを選択"
        If .Show <> -1 Then Exit Sub
        destinationFolder = .SelectedItems(1)
    End With
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim destinationPath As String
    destinationPath = fso.BuildPath(destinationFolder, fso.GetFileName(sourcePath))
    If fso.FileExists(destinationPath) Then
        If MsgBox("同名ファイルがあります。上書きしますか？", vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    End If
    fso.CopyFile sourcePath, destinationPath, True
    MsgBox "コピーしました。" & vbCrLf & destinationPath, vbInformation
End Sub
