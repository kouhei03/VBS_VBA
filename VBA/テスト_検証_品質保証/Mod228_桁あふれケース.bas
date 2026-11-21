Attribute VB_Name = "Mod228_桁あふれケース"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: 桁あふれオーバーフローゼロ除算ケース生成を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「桁あふれオーバーフローゼロ除算ケース生成」を選択して実行する。

' ── 桁あふれオーバーフローゼロ除算ケース生成 ──
Public Sub 桁あふれオーバーフローゼロ除算ケース生成()
    Dim ws As Worksheet: Set ws = GetOrCreateSheet
    ws.Cells.Clear
    ws.Range("A1:C1").Value = Array("ケース", "値", "除数")
    ws.Cells(2, 1).Value = "桁あふれ"
    ws.Cells(2, 2).Value = 999999999999999#
    ws.Cells(3, 1).Value = "ゼロ除算"
    ws.Cells(3, 2).Value = 1
    ws.Cells(3, 3).Value = 0
    ws.Cells(4, 1).Value = "負数"
    ws.Cells(4, 2).Value = -999999999999999#
End Sub

' ── GetOrCreateSheet ──
Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("境界異常ケース")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "境界異常ケース"
    End If
End Function




