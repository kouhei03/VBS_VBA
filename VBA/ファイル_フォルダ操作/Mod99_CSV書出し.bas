Attribute VB_Name = "Mod99_CSV書出し"
Option Explicit

Public Sub 選択範囲CSV書出し()
    Dim destinationFolder As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show <> -1 Then Exit Sub
        destinationFolder = .SelectedItems(1)
    End With
    Dim fileName As String
    fileName = InputBox("CSVファイル名を入力", "CSV書出し", "export.csv")
    If Trim$(fileName) = "" Then Exit Sub
    If LCase$(Right$(fileName, 4)) <> ".csv" Then fileName = fileName & ".csv"

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim temp As Workbook
    Set temp = Workbooks.Add(1)
    Selection.Copy temp.Worksheets(1).Range("A1")
    temp.SaveAs fso.BuildPath(destinationFolder, fileName), xlCSV
    temp.Close False
    MsgBox "CSVを書き出しました。", vbInformation
End Sub
