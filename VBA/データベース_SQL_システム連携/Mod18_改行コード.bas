Attribute VB_Name = "Mod18_改行コード"
Option Explicit

Public Sub 改行コード指定連携ファイル出力()
    Dim mode As String: mode = InputBox("改行コード: 1=CRLF、2=LF", "改行コード", "1")
    If mode <> "1" And mode <> "2" Then Exit Sub
    Dim destination As String
    With Application.FileDialog(msoFileDialogSaveAs)
        If .Show <> -1 Then Exit Sub
        destination = .SelectedItems(1)
    End With
    Dim text As String: text = CStr(Selection.Value)
    If mode = "1" Then
        text = Replace(text, vbLf, vbCrLf)
    Else
        text = Replace(text, vbCrLf, vbLf)
    End If
    Dim stream As Object: Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText text
    stream.SaveToFile destination, 2
    stream.Close
End Sub
