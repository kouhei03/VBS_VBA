Attribute VB_Name = "Mod97_外部ブック読込"
Option Explicit

Public Sub 外部ブック読み取り専用オープン()
    Dim filePath As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .Filters.Clear
        .Filters.Add "Excel", "*.xls;*.xlsx;*.xlsm;*.xlsb"
        If .Show <> -1 Then Exit Sub
        filePath = .SelectedItems(1)
    End With
    Workbooks.Open fileName:=filePath, UpdateLinks:=0, ReadOnly:=True
End Sub
