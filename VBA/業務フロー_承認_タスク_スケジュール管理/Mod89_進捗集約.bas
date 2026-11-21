Attribute VB_Name = "Mod89_進捗集約"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: 複数担当者進捗シート集約全体更新を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「複数担当者進捗シート集約全体更新」を選択して実行する。

' ── 複数担当者進捗シート集約全体更新 ──
Public Sub 複数担当者進捗シート集約全体更新()
    Dim output As Worksheet: Set output = GetOrCreateSheet
    output.Cells.Clear
    Dim outRow As Long: outRow = 1
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "全体進捗" Then
            Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            If lastRow >= 2 Then
                ws.Range("A2:F" & lastRow).Copy output.Cells(outRow, 1)
                outRow = output.Cells(output.Rows.Count, 1).End(xlUp).Row + 1
            End If
        End If
    Next ws
End Sub

' ── GetOrCreateSheet ──
Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("全体進捗")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "全体進捗"
    End If
End Function




