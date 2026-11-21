Attribute VB_Name = "Mod101_JSON_XML"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: JSON_XML読み込みパースを中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、ADO、対象DBドライバまたはOLE DBプロバイダー、DB接続権限。
' ファイル実行例: ExcelのAlt+F8から「JSON_XML読み込みパース」を選択して実行する。

' ── JSON_XML読み込みパース ──
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

' ── WriteXmlNodes ──
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

' ── ReadUtf8 ──
Private Function ReadUtf8(ByVal filePath As String) As String
    Dim stream As Object: Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile filePath
    ReadUtf8 = stream.ReadText
    stream.Close
End Function




