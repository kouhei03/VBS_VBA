Attribute VB_Name = "Mod98_CSV読み込み"
Option Explicit
' ファイル処理概要: ブック・シートを処理を実行する。
' ファイルないコード処理詳細解説: CSV読み込み文字コード対応を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、ADO、対象DBドライバまたはOLE DBプロバイダー、DB接続権限。
' ファイル実行例: ExcelのAlt+F8から「CSV読み込み文字コード対応」を選択して実行する。

' ── CSV読み込み文字コード対応 ──
Public Sub CSV読み込み文字コード対応()
    Dim filePath As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .Filters.Clear
        .Filters.Add "CSV", "*.csv"
        If .Show <> -1 Then Exit Sub
        filePath = .SelectedItems(1)
    End With
    Dim codePage As String
    codePage = InputBox("文字コードを入力: 65001=UTF-8、932=Shift_JIS、1252=ANSI", "CSV読込", "65001")
    If codePage = "65001" Then
        Dim text As String: text = ReadUtf8(filePath)
        Dim data As Variant: data = Split(text, vbCrLf)
        Dim i As Long
        Dim ws As Worksheet: Set ws = ActiveSheet
        For i = LBound(data) To UBound(data)
            ws.Cells(i + 1, 1).Value = data(i)
        Next i
    Else
        Workbooks.OpenText fileName:=filePath, Origin:=CLng(codePage), StartRow:=1, _
                           DataType:=xlDelimited, TextQualifier:=xlTextQualifierDoubleQuote, _
                           Comma:=True, Local:=True
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




