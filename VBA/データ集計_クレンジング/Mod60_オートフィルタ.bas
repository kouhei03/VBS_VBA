Attribute VB_Name = "Mod60_オートフィルタ"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: オートフィルタ設定を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「オートフィルタ設定」を選択して実行する。

' ── オートフィルタ設定 ──
Public Sub オートフィルタ設定()
    Dim ws As Worksheet: Set ws = ActiveSheet
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.UsedRange.AutoFilter
End Sub

' ── オートフィルタ解除 ──
Public Sub オートフィルタ解除()
    If ActiveSheet.AutoFilterMode Then ActiveSheet.AutoFilterMode = False
End Sub




