Attribute VB_Name = "Mod82_ファイル選択"
Option Explicit

Public Sub 単一ファイル選択()
    Dim selected As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        If .Show = -1 Then selected = .SelectedItems(1)
    End With
    If selected <> "" Then MsgBox selected, vbInformation
End Sub

Public Sub 複数ファイル選択()
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = True
        If .Show = -1 Then
            Dim i As Long
            For i = 1 To .SelectedItems.Count
                ActiveSheet.Cells(i, 1).Value = .SelectedItems(i)
            Next i
        End If
    End With
End Sub
