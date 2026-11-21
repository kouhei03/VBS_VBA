Attribute VB_Name = "Mod91_ファイルリネーム"
Option Explicit

Public Sub ファイルリネーム連番日付付与()
    Dim folderPath As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show <> -1 Then Exit Sub
        folderPath = .SelectedItems(1)
    End With
    Dim mode As String
    mode = InputBox("付与方式を入力: 1=連番、2=日付", "ファイルリネーム", "1")
    If mode <> "1" And mode <> "2" Then Exit Sub

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim folder As Object: Set folder = fso.GetFolder(folderPath)
    Dim i As Long: i = 1
    Dim file As Object
    For Each file In folder.Files
        If Left$(file.Name, 2) <> "~$" Then
            Dim suffix As String
            If mode = "1" Then suffix = Format$(i, "000")
            If mode = "2" Then suffix = Format$(file.DateLastModified, "yyyymmdd")
            Dim extension As String: extension = fso.GetExtensionName(file.Name)
            Dim baseName As String: baseName = fso.GetBaseName(file.Name)
            Dim newName As String: newName = suffix & "_" & baseName
            If extension <> "" Then newName = newName & "." & extension
            If fso.BuildPath(folderPath, newName) <> file.Path Then
                fso.MoveFile file.Path, fso.BuildPath(folderPath, newName)
            End If
            i = i + 1
        End If
    Next file
    MsgBox "リネームしました。", vbInformation
End Sub
