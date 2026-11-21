Attribute VB_Name = "Mod100_TSVテキスト"
Option Explicit

Public Sub TSVテキスト読み書き()
    Dim mode As String
    mode = InputBox("処理を入力: 1=読み込み、2=書き出し", "TSV/テキスト", "1")
    If mode = "1" Then
        Dim filePath As String
        With Application.FileDialog(msoFileDialogFilePicker)
            .Filters.Add "Text", "*.txt;*.tsv"
            If .Show <> -1 Then Exit Sub
            filePath = .SelectedItems(1)
        End With
        Dim separator As String: separator = IIf(LCase$(Right$(filePath, 4)) = ".tsv", vbTab, vbLf)
        Dim text As String
        text = ReadText(filePath)
        ActiveSheet.Cells(1, 1).Value = text
    ElseIf mode = "2" Then
        Dim destination As String
        With Application.FileDialog(msoFileDialogSaveAs)
            .Filters.Clear
            .Filters.Add "Text", "*.txt"
            .Filters.Add "TSV", "*.tsv"
            If .Show = -1 Then destination = .SelectedItems(1)
        End With
        If destination <> "" Then WriteText destination, Selection.Text
    End If
End Sub

Private Function ReadText(ByVal filePath As String) As String
    Dim stream As Object: Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile filePath
    ReadText = stream.ReadText
    stream.Close
End Function

Private Sub WriteText(ByVal filePath As String, ByVal text As String)
    Dim stream As Object: Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText text
    stream.SaveToFile filePath, 2
    stream.Close
End Sub
