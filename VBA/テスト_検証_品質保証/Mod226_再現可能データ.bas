Attribute VB_Name = "Mod226_再現可能データ"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: ランダムシード再現可能テストデータ生成を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「ランダムシード再現可能テストデータ生成」を選択して実行する。

' ── ランダムシード再現可能テストデータ生成 ──
Public Sub ランダムシード再現可能テストデータ生成()
    Dim seed As Variant: seed = Application.InputBox("乱数シード", "再現可能データ", 1, , , , , 1)
    Dim count As Variant: count = Application.InputBox("生成件数", "再現可能データ", 100, , , , , 1)
    If seed = False Or count = False Then Exit Sub
    Rnd -1
    Randomize CLng(seed)
    Dim ws As Worksheet: Set ws = GetOrCreateSheet
    ws.Cells.Clear
    Dim i As Long
    For i = 1 To CLng(count)
        ws.Cells(i, 1).Value = i
        ws.Cells(i, 2).Value = Int(Rnd * 100000) / 100
    Next i
End Sub

' ── GetOrCreateSheet ──
Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("再現可能データ")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "再現可能データ"
    End If
End Function




