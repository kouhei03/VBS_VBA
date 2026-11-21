Attribute VB_Name = "Mod135_Word起動"
Option Explicit
' ファイル処理概要: Wordを操作、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: Word文書起動新規作成保存を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、Microsoft Word、COMオートメーション利用可。
' ファイル実行例: ExcelのAlt+F8から「Word文書起動新規作成保存」を選択して実行する。

' ── Word文書起動新規作成保存 ──
Public Sub Word文書起動新規作成保存()
    Dim word As Object
    On Error Resume Next
    Set word = GetObject(, "Word.Application")
    On Error GoTo 0
    If word Is Nothing Then Set word = CreateObject("Word.Application")
    word.Visible = True
    Dim document As Object: Set document = word.Documents.Add
    Dim filePath As String
    With Application.FileDialog(msoFileDialogSaveAs)
        .Filters.Clear
        .Filters.Add "Word", "*.docx"
        If .Show = -1 Then filePath = .SelectedItems(1)
    End With
    If filePath <> "" Then document.SaveAs2 filePath
End Sub




