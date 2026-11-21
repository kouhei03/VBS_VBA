Attribute VB_Name = "Mod138_WordPDF"
Option Explicit

Public Sub Word文書PDF出力()
    Dim word As Object: Set word = CreateObject("Word.Application")
    Dim document As Object: Set document = word.Documents.Open(InputBox("Wordファイルのフルパス", "PDF出力"))
    Dim pdfPath As String: pdfPath = InputBox("PDF保存先フルパス", "PDF出力")
    If pdfPath <> "" Then document.ExportAsFixedFormat pdfPath, 17
    document.Close False
    word.Quit
End Sub
