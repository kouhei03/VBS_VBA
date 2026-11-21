Attribute VB_Name = "Mod94_人員不足検知"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 時間帯日別人員不足過不足検知を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「時間帯日別人員不足過不足検知」を選択して実行する。

' ── 時間帯日別人員不足過不足検知 ──
Public Sub 時間帯日別人員不足過不足検知()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim r As Long
    For r = 2 To lastRow
        If IsNumeric(ws.Cells(r, 3).Value) And IsNumeric(ws.Cells(r, 4).Value) Then
            If CDbl(ws.Cells(r, 3).Value) < CDbl(ws.Cells(r, 4).Value) Then
                ws.Cells(r, 5).Value = "不足"
                ws.Rows(r).Interior.Color = RGB(255, 199, 206)
            Else
                ws.Cells(r, 5).Value = "十分"
            End If
        End If
    Next r
End Sub




