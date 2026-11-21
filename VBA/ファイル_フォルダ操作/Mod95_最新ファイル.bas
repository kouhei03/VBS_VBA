Attribute VB_Name = "Mod95_最新ファイル"
Option Explicit

Public Sub 最新更新ファイル特定とオープン()
    Dim folderPath As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show <> -1 Then Exit Sub
        folderPath = .SelectedItems(1)
    End With
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim latestFile As Object
    Dim file As Object
    For Each file In fso.GetFolder(folderPath).Files
        If latestFile Is Nothing Or file.DateLastModified > latestFile.DateLastModified Then
            Set latestFile = file
        End If
    Next file
    If latestFile Is Nothing Then
        MsgBox "ファイルがありません。", vbInformation
    Else
        MsgBox "最新ファイル: " & latestFile.Name, vbInformation
        Shell "explorer.exe """ & latestFile.Path & """", vbNormalFocus
    End If
End Sub
