Attribute VB_Name = "Mod171_絞り込み検索"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: 複数条件絞り込み検索結果一覧を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「複数条件絞り込み検索結果一覧」を選択して実行する。

' ── 複数条件絞り込み検索結果一覧 ──
Public Sub 複数条件絞り込み検索結果一覧()
    Dim keyword As String: keyword = InputBox("検索文字", "絞り込み検索")
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim output As Worksheet: Set output = GetOrCreateSheet
    output.Cells.Clear
    Dim r As Long, outRow As Long: outRow = 1
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If InStr(1, ws.Rows(r).Text, keyword, vbTextCompare) > 0 Then
            ws.Rows(r).Copy output.Rows(outRow)
            outRow = outRow + 1
        End If
    Next r
End Sub

' ── GetOrCreateSheet ──
Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("検索結果")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "検索結果"
    End If
End Function




