Attribute VB_Name = "Mod160_連携履歴"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 連携処理実行履歴成否ステータス記録を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「連携処理実行履歴成否ステータス記録」を選択して実行する。

' ── 連携処理実行履歴成否ステータス記録 ──
Public Sub 連携処理実行履歴成否ステータス記録()
    Dim ws As Worksheet: Set ws = GetOrCreateSheet(ThisWorkbook, "連携履歴")
    Dim r As Long: r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 2).Value = InputBox("処理名を入力", "連携履歴")
    ws.Cells(r, 3).Value = InputBox("ステータスを入力（成功/警告/失敗）", "連携履歴")
    ws.Cells(r, 4).Value = InputBox("詳細を入力", "連携履歴")
    ws.Cells(r, 1).NumberFormat = "yyyy/mm/dd hh:mm:ss"
    ws.Columns("A:D").AutoFit
End Sub




