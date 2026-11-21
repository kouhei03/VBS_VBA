Attribute VB_Name = "Mod128_操作ログ集計"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: 操作ログアクセスログ読込集計を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、ADO、対象DBドライバまたはOLE DBプロバイダー、DB接続権限。
' ファイル実行例: ExcelのAlt+F8から「操作ログアクセスログ読込集計」を選択して実行する。

' ── 操作ログアクセスログ読込集計 ──
Public Sub 操作ログアクセスログ読込集計()
    Dim filePath As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .Filters.Add "Log/CSV", "*.csv;*.txt;*.log"
        If .Show <> -1 Then Exit Sub
        filePath = .SelectedItems(1)
    End With
    Dim text As String
    text = ReadText(filePath)
    Dim output As Worksheet: Set output = GetOrCreateSheet
    output.Cells.Clear
    output.Cells(1, 1).Value = "ログ内容"
    output.Cells(2, 1).Value = text
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

' ── GetOrCreateSheet ──
Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("ログ集計")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "ログ集計"
    End If
End Function




