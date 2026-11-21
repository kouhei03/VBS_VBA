Attribute VB_Name = "Mod137_Word貼付け"
Option Explicit

Public Sub Excel表グラフWord貼り付け()
    If Selection Is Nothing Then Exit Sub
    Selection.Copy
    Dim word As Object: Set word = CreateObject("Word.Application")
    word.Visible = True
    Dim document As Object: Set document = word.Documents.Add
    document.Content.Paste
    MsgBox "選択範囲を貼り付けました。", vbInformation
End Sub
