Attribute VB_Name = "Mod116_出力先オープン"
Option Explicit

Public Sub 出力先フォルダエクスプローラ表示()
    Dim folderPath As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "表示する出力先フォルダを選択"
        If .Show <> -1 Then Exit Sub
        folderPath = .SelectedItems(1)
    End With
    Shell "explorer.exe """ & folderPath & """", vbNormalFocus
End Sub
