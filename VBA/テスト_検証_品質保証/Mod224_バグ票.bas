Attribute VB_Name = "Mod224_バグ票"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: 不具合票自動起票フォーマット再現条件転記を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「不具合票自動起票フォーマット再現条件転記」を選択して実行する。

' ── 不具合票自動起票フォーマット再現条件転記 ──
Public Sub 不具合票自動起票フォーマット再現条件転記()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim issue As Worksheet: Set issue = GetOrCreateSheet
    issue.Cells.Clear
    issue.Range("A1:B1").Value = Array("項目", "内容")
    issue.Cells(2, 1).Value = "対象"
    issue.Cells(2, 2).Value = Selection.Cells(1, 1).Value
    issue.Cells(3, 1).Value = "再現手順"
    issue.Cells(3, 2).Value = "テスト仕様書の手順に従う"
    issue.Cells(4, 1).Value = "期待値"
    issue.Cells(4, 2).Value = Selection.Cells(1, 2).Value
    issue.Cells(5, 1).Value = "実績値"
    issue.Cells(5, 2).Value = Selection.Cells(1, 3).Value
End Sub

' ── GetOrCreateSheet ──
Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("不具合票")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "不具合票"
    End If
End Function




