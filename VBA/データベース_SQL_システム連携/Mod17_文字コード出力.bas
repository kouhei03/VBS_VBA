Attribute VB_Name = "Mod17_文字コード出力"
Option Explicit

Public Sub 文字コード指定出力ファイル作成()
    Dim destination As String
    With Application.FileDialog(msoFileDialogSaveAs)
        If .Show <> -1 Then Exit Sub
        destination = .SelectedItems(1)
    End With
    Dim charset As String: charset = InputBox("文字コード: 932=Shift-JIS、65001=UTF-8", "文字コード", "65001")
    Dim stream As Object: Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = IIf(charset = "932", "shift_jis", "utf-8")
    stream.Open
    stream.WriteText Selection.Value
    stream.SaveToFile destination, 2
    stream.Close
End Sub
