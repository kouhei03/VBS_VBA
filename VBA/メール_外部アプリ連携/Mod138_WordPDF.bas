Attribute VB_Name = "Mod138_WordPDF"
Option Explicit
' ファイル処理概要: Wordを操作、PDF等へ固定形式出力を実行する。
' ファイルないコード処理詳細解説: Word文書PDF出力を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、Microsoft Word、COMオートメーション利用可。
' ファイル実行例: ExcelのAlt+F8から「Word文書PDF出力」を選択して実行する。

' ── Word文書PDF出力 ──
Public Sub Word文書PDF出力()
    Dim word As Object: Set word = CreateObject("Word.Application")
    Dim document As Object: Set document = word.Documents.Open(InputBox("Wordファイルのフルパス", "PDF出力"))
    Dim pdfPath As String: pdfPath = InputBox("PDF保存先フルパス", "PDF出力")
    If pdfPath <> "" Then document.ExportAsFixedFormat pdfPath, 17
    document.Close False
    word.Quit
End Sub




