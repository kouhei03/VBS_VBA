Attribute VB_Name = "Mod110_古いバックアップ削除"
Option Explicit

Public Sub 古いバックアップファイル自動削除()
    Dim backupFolder As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "バックアップフォルダを選択"
        If .Show <> -1 Then Exit Sub
        backupFolder = .SelectedItems(1)
    End With
    Dim keepDays As Variant
    keepDays = Application.InputBox("保持日数を入力してください", "世代管理", 30, , , , , 1)
    If keepDays = False Or Not IsNumeric(keepDays) Then Exit Sub

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim deletedCount As Long: deletedCount = 0
    Dim cutoff As Date: cutoff = DateAdd("d", -CLng(keepDays), Now)
    Dim file As Object
    For Each file In fso.GetFolder(backupFolder).Files
        If file.DateLastModified < cutoff Then
            file.Delete True
            deletedCount = deletedCount + 1
        End If
    Next file
    MsgBox deletedCount & " 件削除しました。", vbInformation
End Sub
