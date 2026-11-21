Attribute VB_Name = "Mod187_複数シートPDF"
Option Explicit

Public Sub 複数シートまとめてPDF出力()
    Dim outputPath As String
    outputPath = Application.GetSaveAsFilename("帳票一括.pdf", "PDFファイル (*.pdf),*.pdf")
    If outputPath = "False" Then Exit Sub
    Dim names As String
    names = InputBox("出力するシート名をカンマ区切りで入力（空欄で全シート）", "複数シートPDF")
    Dim sheetsToExport As Variant
    If Trim$(names) = "" Then
        ThisWorkbook.ExportAsFixedFormat xlTypePDF, outputPath
    Else
        Dim sheetNames As Variant: sheetNames = Split(names, ",")
        Dim i As Long
        For i = LBound(sheetNames) To UBound(sheetNames)
            sheetNames(i) = Trim$(CStr(sheetNames(i)))
        Next i
        ThisWorkbook.Worksheets(sheetNames).Select
        ActiveSheet.ExportAsFixedFormat xlTypePDF, outputPath
        ActiveSheet.Select
    End If
    MsgBox "PDFを出力しました。", vbInformation
End Sub
