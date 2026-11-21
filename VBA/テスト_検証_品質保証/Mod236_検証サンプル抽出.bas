Attribute VB_Name = "Mod236_検証サンプル抽出"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: 特定条件レコード検証サンプル抽出を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「特定条件レコード検証サンプル抽出」を選択して実行する。

' ── 特定条件レコード検証サンプル抽出 ──
Public Sub 特定条件レコード検証サンプル抽出()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim keyword As String: keyword = InputBox("抽出条件文字", "検証サンプル")
    Dim out As Worksheet: Set out = GetOrCreateSheet
    out.Cells.Clear
    Dim r As Long, outRow As Long: outRow = 1
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If InStr(1, ws.Rows(r).Text, keyword, vbTextCompare) > 0 Then
            ws.Rows(r).Copy out.Rows(outRow)
            outRow = outRow + 1
        End If
    Next r
End Sub

' ── GetOrCreateSheet ──
Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("検証サンプル")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "検証サンプル"
    End If
End Function




