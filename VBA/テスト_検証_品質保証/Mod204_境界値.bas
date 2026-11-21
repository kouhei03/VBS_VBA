Attribute VB_Name = "Mod204_境界値"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: 境界値テストデータ生成を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「境界値テストデータ生成」を選択して実行する。

' ── 境界値テストデータ生成 ──
Public Sub 境界値テストデータ生成()
    Dim minimum As Variant: minimum = Application.InputBox("最小値", "境界値", 0, , , , , 1)
    Dim maximum As Variant: maximum = Application.InputBox("最大値", "境界値", 100, , , , , 1)
    If minimum = False Or maximum = False Then Exit Sub
    Dim ws As Worksheet: Set ws = GetOrCreateSheet
    ws.Cells.Clear
    ws.Range("A1:B1").Value = Array("パターン", "値")
    ws.Range("A2:B8").Value = Array(Array("負数", -1), Array("0", 0), Array("最小値", minimum), _
                                     Array("最小値+1", CDbl(minimum) + 1), Array("最大値-1", CDbl(maximum) - 1), _
                                     Array("最大値", maximum), Array("最大値+1", CDbl(maximum) + 1))
End Sub

' ── GetOrCreateSheet ──
Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("境界値")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "境界値"
    End If
End Function




