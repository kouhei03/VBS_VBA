Attribute VB_Name = "Mod203_異常系データ"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: 異常系空欄文字種不正日付テストデータ生成を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「異常系空欄文字種不正日付テストデータ生成」を選択して実行する。

' ── 異常系空欄文字種不正日付テストデータ生成 ──
Public Sub 異常系空欄文字種不正日付テストデータ生成()
    Dim ws As Worksheet: Set ws = GetOrCreateSheet
    ws.Cells.Clear
    ws.Range("A1:D1").Value = Array("空欄", "文字種誤り", "不正日付", "負数")
    ws.Cells(2, 1).Value = ""
    ws.Cells(2, 2).Value = "１２３４abc"
    ws.Cells(2, 3).Value = "2026/13/40"
    ws.Cells(2, 4).Value = -1
End Sub

' ── GetOrCreateSheet ──
Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("異常系データ")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "異常系データ"
    End If
End Function




