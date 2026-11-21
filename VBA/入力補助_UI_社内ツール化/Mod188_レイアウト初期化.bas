Attribute VB_Name = "Mod188_レイアウト初期化"
Option Explicit
' ファイル処理概要: ブック・シートを処理を実行する。
' ファイルないコード処理詳細解説: 複数シートレイアウト書式ヘッダ一括初期化を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「複数シートレイアウト書式ヘッダ一括初期化」を選択して実行する。

' ── 複数シートレイアウト書式ヘッダ一括初期化 ──
Public Sub 複数シートレイアウト書式ヘッダ一括初期化()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.UsedRange.Font.Name = "Yu Gothic"
        ws.UsedRange.Font.Size = 10
        ws.Rows(1).Font.Bold = True
        ws.Rows(1).Interior.Color = RGB(217, 234, 247)
        ws.Columns.AutoFit
    Next ws
    MsgBox "全シートのレイアウトを初期化しました。", vbInformation
End Sub




