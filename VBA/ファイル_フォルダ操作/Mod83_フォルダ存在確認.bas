Attribute VB_Name = "Mod83_フォルダ存在確認"
Option Explicit

Public Sub フォルダ存在確認と作成()
    Dim folderPath As String
    folderPath = InputBox("確認・作成するフォルダのフルパスを入力してください", "フォルダ作成")
    If Trim$(folderPath) = "" Then Exit Sub

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(folderPath) Then
        MsgBox "フォルダは既に存在します。" & vbCrLf & folderPath, vbInformation
    Else
        If MsgBox("フォルダを作成しますか？" & vbCrLf & folderPath, vbYesNo + vbQuestion) = vbYes Then
            CreateFolderTree fso, folderPath
            MsgBox "フォルダを作成しました。", vbInformation
        End If
    End If
End Sub

Private Sub CreateFolderTree(ByVal fso As Object, ByVal folderPath As String)
    If fso.FolderExists(folderPath) Then Exit Sub
    Dim parentPath As String
    parentPath = fso.GetParentFolderName(folderPath)
    If parentPath <> "" And Not fso.FolderExists(parentPath) Then CreateFolderTree fso, parentPath
    fso.CreateFolder folderPath
End Sub
