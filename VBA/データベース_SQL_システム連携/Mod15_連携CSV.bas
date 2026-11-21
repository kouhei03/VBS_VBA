Attribute VB_Name = "Mod15_連携CSV"
Option Explicit

Public Sub 基幹システム取込CSV一括生成()
    Dim destination As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show <> -1 Then Exit Sub
        destination = .SelectedItems(1)
    End With
    Dim temp As Workbook: Set temp = Workbooks.Add(1)
    ActiveSheet.UsedRange.Copy temp.Worksheets(1).Range("A1")
    temp.SaveAs destination & "\import_" & Format$(Now, "yyyymmdd_hhnnss") & ".csv", xlCSV
    temp.Close False
    MsgBox "連携CSVを生成しました。", vbInformation
End Sub
