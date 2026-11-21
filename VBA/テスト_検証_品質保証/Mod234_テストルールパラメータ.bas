Attribute VB_Name = "Mod234_テストルールパラメータ"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: テストデータルール件数発生率対象項目パラメータ化を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「テストデータルール件数発生率対象項目パラメータ化」を選択して実行する。

' ── テストデータルール件数発生率対象項目パラメータ化 ──
Public Sub テストデータルール件数発生率対象項目パラメータ化()
    Dim ws As Worksheet: Set ws = GetOrCreateSheet
    ws.Cells.Clear
    ws.Range("A1:B1").Value = Array("パラメータ", "値")
    ws.Cells(2, 1).Value = "件数"
    ws.Cells(2, 2).Value = InputBox("件数", "テストルール", "100")
    ws.Cells(3, 1).Value = "異常発生率"
    ws.Cells(3, 2).Value = InputBox("異常発生率（%）", "テストルール", "10")
    ws.Cells(4, 1).Value = "対象項目"
    ws.Cells(4, 2).Value = InputBox("対象項目", "テストルール")
End Sub

' ── GetOrCreateSheet ──
Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("テストルール")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "テストルール"
    End If
End Function




