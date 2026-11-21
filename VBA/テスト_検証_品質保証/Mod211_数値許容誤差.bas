Attribute VB_Name = "Mod211_数値許容誤差"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 数値許容誤差丸め桁数検証を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「数値許容誤差丸め桁数検証」を選択して実行する。

' ── 数値許容誤差丸め桁数検証 ──
Public Sub 数値許容誤差丸め桁数検証()
    Dim tolerance As Variant: tolerance = Application.InputBox("許容誤差", "数値検証", 0.01, , , , , 1)
    If tolerance = False Or Not IsNumeric(tolerance) Then Exit Sub
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim r As Long
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If IsNumeric(ws.Cells(r, 1).Value) And IsNumeric(ws.Cells(r, 2).Value) Then
            ws.Cells(r, 3).Value = IIf(Abs(CDbl(ws.Cells(r, 1).Value) - CDbl(ws.Cells(r, 2).Value)) _
                                       <= CDbl(tolerance), "OK", "NG")
        End If
    Next r
End Sub




