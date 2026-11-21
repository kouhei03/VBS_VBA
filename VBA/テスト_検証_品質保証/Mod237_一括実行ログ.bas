Attribute VB_Name = "Mod237_一括実行ログ"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: 複数パターン一括実行結果ログ集約を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「複数パターン一括実行結果ログ集約」を選択して実行する。

' ── 複数パターン一括実行結果ログ集約 ──
Public Sub 複数パターン一括実行結果ログ集約()
    Dim log As Worksheet: Set log = GetOrCreateSheet
    log.Cells.Clear
    log.Range("A1:C1").Value = Array("実行日時", "パターン", "結果")
    Dim count As Variant: count = Application.InputBox("実行パターン数", "一括実行", 10, , , , , 1)
    If count = False Then Exit Sub
    Dim i As Long
    For i = 1 To CLng(count)
        log.Cells(i + 1, 1).Value = Now
        log.Cells(i + 1, 2).Value = "Pattern-" & i
        log.Cells(i + 1, 3).Value = "完了"
    Next i
End Sub

' ── GetOrCreateSheet ──
Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("一括実行ログ")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "一括実行ログ"
    End If
End Function




