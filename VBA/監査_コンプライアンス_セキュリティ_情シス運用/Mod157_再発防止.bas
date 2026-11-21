Attribute VB_Name = "Mod157_再発防止"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: インシデント再発防止策実施状況追跡を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「インシデント再発防止策実施状況追跡」を選択して実行する。

' ── インシデント再発防止策実施状況追跡 ──
Public Sub インシデント再発防止策実施状況追跡()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim r As Long
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If CStr(ws.Cells(r, 3).Value) <> "完了" And IsDate(ws.Cells(r, 4).Value) _
           And CDate(ws.Cells(r, 4).Value) < Date Then
            ws.Rows(r).Interior.Color = RGB(255, 199, 206)
        End If
    Next r
End Sub




