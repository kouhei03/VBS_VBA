Attribute VB_Name = "Mod135_ネットワーク棚卸"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: ファイアウォールセキュリティグループルーティング棚卸を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「ファイアウォールセキュリティグループルーティング棚卸」を選択して実行する。

' ── ファイアウォールセキュリティグループルーティング棚卸 ──
Public Sub ファイアウォールセキュリティグループルーティング棚卸()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim output As Worksheet: Set output = GetOrCreateSheet
    output.Cells.Clear
    ws.UsedRange.Copy output.Range("A1")
    output.Cells(1, ws.UsedRange.Columns.Count + 1).Value = "棚卸判定"
End Sub

' ── GetOrCreateSheet ──
Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("ネットワーク棚卸")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "ネットワーク棚卸"
    End If
End Function




