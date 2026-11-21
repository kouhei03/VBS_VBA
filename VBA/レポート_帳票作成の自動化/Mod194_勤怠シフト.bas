Attribute VB_Name = "Mod194_勤怠シフト"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 勤怠集計シフト表整形出力を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「勤怠集計シフト表整形出力」を選択して実行する。

' ── 勤怠集計シフト表整形出力 ──
Public Sub 勤怠集計シフト表整形出力()
    Dim ws As Worksheet: Set ws = ActiveSheet
    ws.UsedRange.Columns.AutoFit
    ws.UsedRange.Rows.AutoFit
    ws.Rows(1).Font.Bold = True
    ws.Rows(1).Interior.Color = RGB(47, 85, 151)
    ws.Rows(1).Font.Color = RGB(255, 255, 255)
    ws.PageSetup.Orientation = xlLandscape
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.FitToPagesTall = 0
End Sub




