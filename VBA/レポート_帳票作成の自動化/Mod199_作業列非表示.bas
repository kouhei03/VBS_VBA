Attribute VB_Name = "Mod199_作業列非表示"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 印刷不要作業列メモ非表示を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「印刷不要作業列メモ非表示」を選択して実行する。

' ── 印刷不要作業列メモ非表示 ──
Public Sub 印刷不要作業列メモ非表示()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim keyword As String: keyword = InputBox("非表示にする列見出しキーワード（例: メモ）", "作業列非表示")
    If Trim$(keyword) = "" Then Exit Sub
    Dim lastColumn As Long: lastColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Dim c As Long
    For c = 1 To lastColumn
        If InStr(1, CStr(ws.Cells(1, c).Value), keyword, vbTextCompare) > 0 Then
            ws.Columns(c).Hidden = True
        End If
    Next c
End Sub




