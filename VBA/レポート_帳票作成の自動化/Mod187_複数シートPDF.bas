Attribute VB_Name = "Mod187_複数シートPDF"
Option Explicit
' ファイル処理概要: PDF等へ固定形式出力、ブック・シートを処理を実行する。
' ファイルないコード処理詳細解説: 複数シートまとめてPDF出力を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「複数シートまとめてPDF出力」を選択して実行する。

' ── 複数シートまとめてPDF出力 ──
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




