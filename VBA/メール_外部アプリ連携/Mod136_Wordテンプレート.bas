Attribute VB_Name = "Mod136_Wordテンプレート"
Option Explicit

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
