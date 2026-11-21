Attribute VB_Name = "Mod127_複数添付"
Option Explicit

Public Sub 複数ファイル添付()
    Dim outlook As Object: Set outlook = GetOutlook()
    Dim mail As Object: Set mail = outlook.CreateItem(0)
    mail.To = InputBox("宛先を入力してください", "添付メール")
    mail.Subject = InputBox("件名を入力してください", "添付メール")
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = True
        If .Show = -1 Then
            Dim i As Long
            For i = 1 To .SelectedItems.Count
                mail.Attachments.Add .SelectedItems(i)
            Next i
        End If
    End With
    mail.Save
End Sub
