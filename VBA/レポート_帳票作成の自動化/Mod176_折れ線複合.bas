Attribute VB_Name = "Mod176_折れ線複合"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 折れ線複合グラフ自動生成を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「折れ線複合グラフ自動生成」を選択して実行する。

' ── 折れ線複合グラフ自動生成 ──
Public Sub 折れ線複合グラフ自動生成()
    Dim dataRange As Range
    Set dataRange = Application.InputBox("グラフ化する範囲を選択", "折れ線グラフ", Selection.Address, , , , , 8)
    Dim chart As Chart
    Set chart = ActiveSheet.Shapes.AddChart2(227, xlLineMarkers).Chart
    chart.SetSourceData dataRange
    chart.HasTitle = True
    chart.ChartTitle.Text = "推移グラフ"
End Sub




