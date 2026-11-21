Attribute VB_Name = "Mod40_暗号化圧縮補助"
Option Explicit

Public Sub 外部連携データ暗号化圧縮Zip前処理補助()
    Dim sourceFolder As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "圧縮する連携ファイルがあるフォルダを選択"
        If .Show <> -1 Then Exit Sub
        sourceFolder = .SelectedItems(1)
    End With
    Dim zipPath As String
    zipPath = Application.GetSaveAsFilename("連携データ.zip", "Zipファイル (*.zip),*.zip")
    If zipPath = "False" Then Exit Sub
    Dim shell As Object: Set shell = CreateObject("Shell.Application")
    Dim zipFolder As Object
    CreateEmptyZip zipPath
    Set zipFolder = shell.NameSpace(zipPath)
    zipFolder.CopyHere shell.NameSpace(sourceFolder).Items, 16
    MsgBox "圧縮処理を開始しました。" & vbCrLf & zipPath, vbInformation
End Sub

Private Sub CreateEmptyZip(ByVal zipPath As String)
    Dim fileNumber As Integer: fileNumber = FreeFile
    Open zipPath For Output As #fileNumber
    Print #fileNumber, "PK" & Chr$(5) & Chr$(6) & String$(18, Chr$(0))
    Close #fileNumber
End Sub
