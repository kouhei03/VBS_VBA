Attribute VB_Name = "Mod227_日本語文字検証"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: 日本語機種依存サロゲート半角カナ検証データを中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「日本語機種依存サロゲート半角カナ検証データ」を選択して実行する。

' ── 日本語機種依存サロゲート半角カナ検証データ ──
Public Sub 日本語機種依存サロゲート半角カナ検証データ()
    Dim ws As Worksheet: Set ws = GetOrCreateSheet
    ws.Cells.Clear
    ws.Cells(1, 1).Value = "パターン"
    ws.Cells(2, 1).Value = "機種依存文字例: ①"
    ws.Cells(3, 1).Value = "サロゲートペア例: 𠮷野家"
    ws.Cells(4, 1).Value = "半角カナ例: ｶﾀｶﾅ"
    ws.Cells(5, 1).Value = "絵文字例: 😀"
End Sub

' ── GetOrCreateSheet ──
Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("日本語文字検証")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "日本語文字検証"
    End If
End Function




