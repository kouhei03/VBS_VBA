Attribute VB_Name = "Mod74_ランキング"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: ランキング順位を付与を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「ランキング順位を付与」を選択して実行する。

' ── ランキング順位を付与 ──
Public Sub ランキング順位を付与()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim valueCol As Long: valueCol = CLng(InputBox("ランキング対象列番号", "ランキング", "3"))
    Dim rankCol As Long: rankCol = GetLastCol(ws) + 1
    ws.Cells(1, rankCol).Value = "ランキング"
    Dim r As Long
    For r = 2 To GetLastRow(ws)
        If IsNumeric(ws.Cells(r, valueCol).Value) Then
            ws.Cells(r, rankCol).Formula = "=RANK(" & ws.Cells(r, valueCol).Address(False, False) & "," & _
                                           ws.Range(ws.Cells(2, valueCol), ws.Cells(GetLastRow(ws), valueCol)).Address(True, True) & ",0)"
        End If
    Next r
End Sub




