Attribute VB_Name = "Mod159_インポートファイル"
Option Explicit

Public Sub 外部システム連携インポートファイル自動生成()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim destinationFolder As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show <> -1 Then Exit Sub
        destinationFolder = .SelectedItems(1)
    End With
    Dim fileName As String: fileName = "import_" & Format$(Now, "yyyymmdd_hhnnss") & ".csv"
    Dim temp As Workbook: Set temp = Workbooks.Add(1)
    ws.UsedRange.Copy temp.Worksheets(1).Range("A1")
    temp.SaveAs destinationFolder & "\" & fileName, xlCSVUTF8
    temp.Close False
    MsgBox "インポートファイルを生成しました。" & vbCrLf & destinationFolder & "\" & fileName, vbInformation
End Sub
