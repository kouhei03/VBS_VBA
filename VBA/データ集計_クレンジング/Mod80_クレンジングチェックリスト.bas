Attribute VB_Name = "Mod80_クレンジングチェックリスト"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: クレンジング前後チェックリスト作成を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「クレンジング前後チェックリスト作成」を選択して実行する。

' ── クレンジング前後チェックリスト作成 ──
Public Sub クレンジング前後チェックリスト作成()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim out As Worksheet: Set out = GetOrCreateSheet(ws.Parent, "クレンジングチェック")
    out.Cells.Clear
    out.Cells(1, 1).Value = "項目"
    out.Cells(1, 2).Value = "件数"
    out.Cells(1, 3).Value = "合計"
    out.Cells(2, 1).Value = "全行"
    out.Cells(2, 2).Value = GetLastRow(ws) - 1
    out.Cells(3, 1).Value = "空白セル"
    out.Cells(3, 2).Value = Application.CountBlank(ws.UsedRange)
    out.Cells(4, 1).Value = "数値合計"
    out.Cells(4, 3).Value = Application.WorksheetFunction.Sum(ws.UsedRange)
End Sub




