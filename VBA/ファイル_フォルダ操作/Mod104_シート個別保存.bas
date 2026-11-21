Attribute VB_Name = "Mod104_シート個別保存"
Option Explicit

Public Sub 複数シートを個別ブックへ分割保存()
    Dim destinationFolder As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show <> -1 Then Exit Sub
        destinationFolder = .SelectedItems(1)
    End With
    Application.DisplayAlerts = False
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        Dim temp As Workbook
        Set temp = Workbooks.Add(1)
        ws.Copy Before:=temp.Worksheets(1)
        temp.Worksheets(1).Delete
        temp.SaveAs destinationFolder & "\" & SafeSheetName(ws.Name) & ".xlsx", xlOpenXMLWorkbook
        temp.Close False
    Next ws
    Application.DisplayAlerts = True
    MsgBox "シートを個別ブックに分割保存しました。", vbInformation
End Sub

Private Function SafeSheetName(ByVal name As String) As String
    Dim invalid As Variant: invalid = Array("\", "/", ":", "*", "?", Chr$(34), "<", ">", "|")
    Dim i As Long
    For i = LBound(invalid) To UBound(invalid)
        name = Replace(name, invalid(i), "_")
    Next i
    SafeSheetName = name
End Function
