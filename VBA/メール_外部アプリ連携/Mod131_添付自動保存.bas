Attribute VB_Name = "Mod131_添付自動保存"
Option Explicit

Public Sub 受信メール添付ファイル自動保存()
    Dim destination As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show <> -1 Then Exit Sub
        destination = .SelectedItems(1)
    End With
    Dim outlook As Object: Set outlook = GetOutlook()
    Dim inbox As Object: Set inbox = outlook.GetNamespace("MAPI").GetDefaultFolder(6)
    Dim count As Long
    Dim item As Object
    For Each item In inbox.Items
        Dim attachment As Object
        For Each attachment In item.Attachments
            attachment.SaveAsFile destination & "\" & attachment.FileName
            count = count + 1
        Next attachment
    Next item
    MsgBox count & " 件保存しました。", vbInformation
End Sub
