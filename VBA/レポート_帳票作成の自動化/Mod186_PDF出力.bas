Attribute VB_Name = "Mod186_PDF出力"
Option Explicit
' ファイル処理概要: PDF等へ固定形式出力を実行する。
' ファイルないコード処理詳細解説: 帳票シートPDF出力を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「帳票シートPDF出力」を選択して実行する。

' ── 帳票シートPDF出力 ──
Public Sub 帳票シートPDF出力()
    Dim outputPath As String
    outputPath = Application.GetSaveAsFilename(ActiveSheet.Name & ".pdf", "PDFファイル (*.pdf),*.pdf")
    If outputPath = "False" Then Exit Sub
    ActiveSheet.ExportAsFixedFormat xlTypePDF, outputPath
    MsgBox "PDFを出力しました。" & vbCrLf & outputPath, vbInformation
End Sub




