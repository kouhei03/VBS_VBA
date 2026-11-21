Attribute VB_Name = "Mod93_フォルダ操作"
Option Explicit

Public Sub フォルダコピー移動削除()
    Dim sourceFolder As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "対象フォルダを選択"
        If .Show <> -1 Then Exit Sub
        sourceFolder = .SelectedItems(1)
    End With
    Dim operation As String
    operation = InputBox("操作を入力: 1=コピー、2=移動、3=削除", "フォルダ操作")

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If operation = "1" Or operation = "2" Then
        Dim destinationParent As String
        With Application.FileDialog(msoFileDialogFolderPicker)
            .Title = "移動先・コピー先の親フォルダを選択"
            If .Show <> -1 Then Exit Sub
            destinationParent = .SelectedItems(1)
        End With
        Dim destinationPath As String
        destinationPath = fso.BuildPath(destinationParent, fso.GetFolder(sourceFolder).Name)
        If operation = "1" Then fso.CopyFolder sourceFolder, destinationPath, True
        If operation = "2" Then fso.MoveFolder sourceFolder, destinationPath
    ElseIf operation = "3" Then
        If MsgBox("フォルダを完全削除しますか？" & vbCrLf & sourceFolder, _
                  vbYesNo + vbCritical) = vbYes Then fso.DeleteFolder sourceFolder, True
    End If
End Sub
