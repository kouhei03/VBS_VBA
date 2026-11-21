Attribute VB_Name = "Mod169_フォント統一"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: フォント種類サイズ太字文字色統一を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「フォント種類サイズ太字文字色統一」を選択して実行する。

' ── フォント種類サイズ太字文字色統一 ──
Public Sub フォント種類サイズ太字文字色統一()
    Dim ws As Worksheet: Set ws = ActiveSheet
    ws.UsedRange.Font.Name = "Yu Gothic"
    ws.UsedRange.Font.Size = 10
    ws.UsedRange.Font.Color = RGB(32, 32, 32)
    ws.Rows(1).Font.Bold = True
    ws.Rows(1).Font.Color = RGB(255, 255, 255)
    ws.Rows(1).Interior.Color = RGB(47, 85, 151)
End Sub




