Attribute VB_Name = "Mod81_フォルダ選択"
Option Explicit

Public Sub フォルダ選択パス取得()
    Dim folderPath As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "フォルダを選択してください"
        If .Show = -1 Then
            folderPath = .SelectedItems(1)
            MsgBox "選択フォルダ: " & vbCrLf & folderPath, vbInformation
        End If
    End With
End Sub
