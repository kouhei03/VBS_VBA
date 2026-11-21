Attribute VB_Name = "Mod177_円ドーナツ"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 円ドーナツグラフ自動作成を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「円ドーナツグラフ自動作成」を選択して実行する。

' ── 円ドーナツグラフ自動作成 ──
Public Sub 円ドーナツグラフ自動作成()
    Dim dataRange As Range
    Set dataRange = Application.InputBox("グラフ化する範囲を選択", "円グラフ", Selection.Address, , , , , 8)
    Dim mode As String: mode = InputBox("種類: 1=円、2=ドーナツ", "円グラフ", "1")
    Dim chartType As XlChartType
    If mode = "1" Then chartType = xlPie Else chartType = xlDoughnut
    ActiveSheet.Shapes.AddChart2(0, chartType).Chart.SetSourceData dataRange
End Sub




