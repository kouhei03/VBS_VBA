Attribute VB_Name = "Mod179_複数グラフ更新"
Option Explicit
' ファイル処理概要: ブック・シートを処理を実行する。
' ファイルないコード処理詳細解説: 複数シートグラフ一括更新レイアウト調整を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「複数シートグラフ一括更新レイアウト調整」を選択して実行する。

' ── 複数シートグラフ一括更新レイアウト調整 ──
Public Sub 複数シートグラフ一括更新レイアウト調整()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        Dim chart As ChartObject
        For Each chart In ws.ChartObjects
            chart.Chart.Refresh
            chart.Width = 600
            chart.Height = 360
        Next chart
    Next ws
    MsgBox "全シートのグラフを更新しました。", vbInformation
End Sub




