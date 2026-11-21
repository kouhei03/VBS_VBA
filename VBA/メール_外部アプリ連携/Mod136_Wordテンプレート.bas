Attribute VB_Name = "Mod136_Wordテンプレート"
Option Explicit
' ファイル処理概要: Wordを操作を実行する。
' ファイルないコード処理詳細解説: Wordテンプレートテキストブックマーク差し込みを中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、Microsoft Word、COMオートメーション利用可。
' ファイル実行例: ExcelのAlt+F8から「Wordテンプレートテキストブックマーク差し込み」を選択して実行する。

' ── Wordテンプレートテキストブックマーク差し込み ──
Public Sub Wordテンプレートテキストブックマーク差し込み()
    Dim templatePath As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .Filters.Add "Word", "*.docx;*.dotx"
        If .Show <> -1 Then Exit Sub
        templatePath = .SelectedItems(1)
    End With
    Dim word As Object: Set word = CreateObject("Word.Application")
    word.Visible = True
    Dim document As Object: Set document = word.Documents.Open(templatePath)
    Dim bookmarkName As String: bookmarkName = InputBox("ブックマーク名", "Word差し込み")
    Dim bookmarkText As String: bookmarkText = InputBox("挿入するテキスト", "Word差し込み")
    If document.Bookmarks.Exists(bookmarkName) Then
        document.Bookmarks(bookmarkName).Range.Text = bookmarkText
    Else
        MsgBox "ブックマークが見つかりません。" & vbCrLf & bookmarkName, vbExclamation
    End If
End Sub




