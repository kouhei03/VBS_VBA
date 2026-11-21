Attribute VB_Name = "Mod218_帳票検証"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 帳票項目欠損合計値一致検証を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「帳票項目欠損合計値一致検証」を選択して実行する。

' ── 帳票項目欠損合計値一致検証 ──
Public Sub 帳票項目欠損合計値一致検証()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim r As Long
    For r = 1 To ws.UsedRange.Rows.Count
        If Application.CountBlank(ws.Rows(r)) > 0 Then
            ws.Cells(r, 1).Interior.Color = RGB(255, 235, 156)
        End If
    Next r
End Sub




