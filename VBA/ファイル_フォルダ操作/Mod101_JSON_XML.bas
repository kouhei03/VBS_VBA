Attribute VB_Name = "Mod101_JSON_XML"
Option Explicit

Public Sub JSON_XML読み込みパース()
    Dim filePath As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .Filters.Clear
        .Filters.Add "JSON/XML", "*.json;*.xml"
        If .Show <> -1 Then Exit Sub
        filePath = .SelectedItems(1)
    End With

    Dim extension As String
    extension = LCase$(Right$(filePath, 5))
    Dim ws As Worksheet: Set ws = ActiveSheet
    If extension Like "*.json" Then
        ws.Cells(1, 1).Value = "JSONテキスト"
        ws.Cells(2, 1).Value = ReadUtf8(filePath)
    Else
        Dim xml As Object: Set xml = CreateObject("MSXML2.DOMDocument")
        xml.Async = False
        If Not xml.Load(filePath) Then
            MsgBox "XMLの解析に失敗しました。" & vbCrLf & xml.ParseError.Reason, vbCritical
            Exit Sub
        End If
        ws.Cells(1, 1).Value = "ノード名"
        ws.Cells(1, 2).Value = "値"
        WriteXmlNodes xml.DocumentElement, ws, 2
    End If
End Sub

Private Sub WriteXmlNodes(ByVal node As Object, ByVal ws As Worksheet, ByRef r As Long)
    Dim child As Object
    If node.ChildNodes.Length > 0 And node.FirstChild.NodeType = 1 Then
        For Each child In node.ChildNodes
            WriteXmlNodes child, ws, r
        Next child
    Else
        ws.Cells(r, 1).Value = node.BaseName
        ws.Cells(r, 2).Value = node.Text
        r = r + 1
    End If
End Sub

Private Function ReadUtf8(ByVal filePath As String) As String
    Dim stream As Object: Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile filePath
    ReadUtf8 = stream.ReadText
    stream.Close
End Function
