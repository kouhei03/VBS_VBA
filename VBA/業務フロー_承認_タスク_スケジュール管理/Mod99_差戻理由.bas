Attribute VB_Name = "Mod99_差戻理由"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: 差戻理由修正コメントログ集計を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「差戻理由修正コメントログ集計」を選択して実行する。

' ── 差戻理由修正コメントログ集計 ──
Public Sub 差戻理由修正コメントログ集計()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim log As Worksheet: Set log = GetOrCreateSheet
    Dim r As Long: r = log.Cells(log.Rows.Count, 1).End(xlUp).Row + 1
    log.Cells(r, 1).Value = Now
    log.Cells(r, 2).Value = CStr(ws.Cells(Selection.Row, 1).Value)
    log.Cells(r, 3).Value = InputBox("差し戻し理由・修正コメント", "差戻ログ")
End Sub

' ── GetOrCreateSheet ──
Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("差戻ログ")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "差戻ログ"
    End If
End Function




