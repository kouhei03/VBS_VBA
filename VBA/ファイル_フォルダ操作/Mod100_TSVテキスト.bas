Attribute VB_Name = "Mod100_TSVテキスト"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: TSVテキスト読み書きを中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、ADO、対象DBドライバまたはOLE DBプロバイダー、DB接続権限。
' ファイル実行例: ExcelのAlt+F8から「TSVテキスト読み書き」を選択して実行する。

' ── TSVテキスト読み書き ──
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

' ── ReadText ──
Private Function ReadText(ByVal filePath As String) As String
    Dim stream As Object: Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile filePath
    ReadText = stream.ReadText
    stream.Close
End Function

' ── WriteText ──
Private Sub WriteText(ByVal filePath As String, ByVal text As String)
    Dim stream As Object: Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText text
    stream.SaveToFile filePath, 2
    stream.Close
End Sub




