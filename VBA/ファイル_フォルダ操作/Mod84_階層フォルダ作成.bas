Attribute VB_Name = "Mod84_階層フォルダ作成"
Option Explicit

Public Sub 階層フォルダ一括作成()
    Dim basePath As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "作成先の親フォルダを選択"
        If .Show <> -1 Then Exit Sub
        basePath = .SelectedItems(1)
    End With

    Dim names As String
    names = InputBox("作成する相対パスを1行ずつ入力（例: 2026\01\A社）", "階層フォルダ作成")
    If Trim$(names) = "" Then Exit Sub

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim lines As Variant: lines = Split(names, vbCrLf)
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        Dim relativePath As String: relativePath = Trim$(CStr(lines(i)))
        If relativePath <> "" Then
            Dim targetPath As String: targetPath = fso.BuildPath(basePath, relativePath)
            CreateFolderTree fso, targetPath
        End If
    Next i
    MsgBox "階層フォルダを作成しました。", vbInformation
End Sub

Private Sub CreateFolderTree(ByVal fso As Object, ByVal folderPath As String)
    If fso.FolderExists(folderPath) Then Exit Sub
    Dim parentPath As String: parentPath = fso.GetParentFolderName(folderPath)
    If parentPath <> "" And Not fso.FolderExists(parentPath) Then CreateFolderTree fso, parentPath
    fso.CreateFolder folderPath
End Sub
