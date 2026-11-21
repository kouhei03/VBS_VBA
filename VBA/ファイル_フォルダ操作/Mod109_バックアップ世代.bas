Attribute VB_Name = "Mod109_バックアップ世代"
Option Explicit

Public Sub バックアップフォルダへ世代管理保存()
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(ThisWorkbook.FullName) Then
        MsgBox "保存済みブックに対して実行してください。", vbExclamation
        Exit Sub
    End If
    Dim backupParent As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "バックアップ先フォルダを選択"
        If .Show <> -1 Then Exit Sub
        backupParent = .SelectedItems(1)
    End With
    Dim generationFolder As String
    generationFolder = fso.BuildPath(backupParent, Format$(Now, "yyyymmdd_hhnnss"))
    fso.CreateFolder generationFolder
    fso.CopyFile ThisWorkbook.FullName, fso.BuildPath(generationFolder, fso.GetFileName(ThisWorkbook.FullName)), True
    MsgBox "世代バックアップを保存しました。" & vbCrLf & generationFolder, vbInformation
End Sub
