Attribute VB_Name = "Mod167_ゼブラ"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 交互行ゼブラカラー適用を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「交互行ゼブラカラー適用」を選択して実行する。

' ── 交互行ゼブラカラー適用 ──
Public Sub 交互行ゼブラカラー適用()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim lastColumn As Long: lastColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Dim r As Long
    For r = 2 To lastRow
        If r Mod 2 = 0 Then ws.Range(ws.Cells(r, 1), ws.Cells(r, lastColumn)).Interior.Color = RGB(242, 242, 242)
    Next r
End Sub




