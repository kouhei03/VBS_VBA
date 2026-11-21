Attribute VB_Name = "Mod120_処理完了アーカイブ"
Option Explicit

Public Sub 処理完了後元データをアーカイブ移動()
    Dim sourceFolder As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "元データフォルダを選択"
        If .Show <> -1 Then Exit Sub
        sourceFolder = .SelectedItems(1)
    End With
    Dim archiveRoot As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "アーカイブ先フォルダを選択"
        If .Show <> -1 Then Exit Sub
        archiveRoot = .SelectedItems(1)
    End With

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim archiveFolder As String
    archiveFolder = fso.BuildPath(archiveRoot, "Archive_" & Format$(Now, "yyyymmdd_hhnnss"))
    fso.CreateFolder archiveFolder
    If MsgBox("処理完了後の元データをアーカイブへ移動しますか？", _
              vbYesNo + vbQuestion) <> vbYes Then
        fso.DeleteFolder archiveFolder
        Exit Sub
    End If
    Dim folder As Object: Set folder = fso.GetFolder(sourceFolder)
    Dim file As Object
    For Each file In folder.Files
        fso.MoveFile file.Path, fso.BuildPath(archiveFolder, file.Name)
    Next file
    MsgBox "元データをアーカイブへ移動しました。" & vbCrLf & archiveFolder, vbInformation
End Sub
