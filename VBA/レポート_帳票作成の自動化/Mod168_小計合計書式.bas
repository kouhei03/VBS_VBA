Attribute VB_Name = "Mod168_小計合計書式"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 合計小計行の太線二重線塗りつぶしを中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「合計小計行の太線二重線塗りつぶし」を選択して実行する。

' ── 合計小計行の太線二重線塗りつぶし ──
Public Sub 合計小計行の太線二重線塗りつぶし()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim r As Long
    For r = 1 To lastRow
        If InStr(CStr(ws.Cells(r, 1).Value), "合計") > 0 Or _
           InStr(CStr(ws.Cells(r, 1).Value), "小計") > 0 Then
            ws.Rows(r).Font.Bold = True
            ws.Rows(r).Borders(xlTop).LineStyle = xlDouble
            ws.Rows(r).Interior.Color = RGB(217, 234, 247)
        End If
    Next r
End Sub




