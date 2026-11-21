Attribute VB_Name = "Mod100_承認待ち滞留"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 承認待ち滞留案件滞留日数カウントを中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「承認待ち滞留案件滞留日数カウント」を選択して実行する。

' ── 承認待ち滞留案件滞留日数カウント ──
Public Sub 承認待ち滞留案件滞留日数カウント()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    ws.Cells(1, 4).Value = "滞留日数"
    Dim r As Long
    For r = 2 To lastRow
        If CStr(ws.Cells(r, 3).Value) = "承認待ち" And IsDate(ws.Cells(r, 2).Value) Then
            ws.Cells(r, 4).Value = DateDiff("d", CDate(ws.Cells(r, 2).Value), Date)
        End If
    Next r
End Sub




