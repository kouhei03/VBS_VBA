Attribute VB_Name = "Mod186_PDF出力"
Option Explicit

Public Sub 帳票シートPDF出力()
    Dim outputPath As String
    outputPath = Application.GetSaveAsFilename(ActiveSheet.Name & ".pdf", "PDFファイル (*.pdf),*.pdf")
    If outputPath = "False" Then Exit Sub
    ActiveSheet.ExportAsFixedFormat xlTypePDF, outputPath
    MsgBox "PDFを出力しました。" & vbCrLf & outputPath, vbInformation
End Sub
