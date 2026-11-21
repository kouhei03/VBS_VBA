Attribute VB_Name = "Mod229_拠点通貨税率"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: 複数拠点通貨税率テストデータ生成を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「複数拠点通貨税率テストデータ生成」を選択して実行する。

' ── 複数拠点通貨税率テストデータ生成 ──
Public Sub 複数拠点通貨税率テストデータ生成()
    Dim branches As Variant: branches = Array("東京", "大阪", "福岡")
    Dim currencies As Variant: currencies = Array("JPY", "USD", "EUR")
    Dim taxRates As Variant: taxRates = Array(0, 8, 10)
    Dim ws As Worksheet: Set ws = GetOrCreateSheet
    ws.Cells.Clear
    ws.Range("A1:C1").Value = Array("拠点", "通貨", "税率")
    Dim i As Long, r As Long: r = 2
    For i = LBound(branches) To UBound(branches)
        ws.Cells(r, 1).Value = branches(i)
        ws.Cells(r, 2).Value = currencies(i)
        ws.Cells(r, 3).Value = taxRates(i)
        r = r + 1
    Next i
End Sub

' ── GetOrCreateSheet ──
Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("拠点通貨税率")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "拠点通貨税率"
    End If
End Function




