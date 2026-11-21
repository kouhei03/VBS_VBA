Attribute VB_Name = "Mod175_縦棒横棒"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 縦棒横棒グラフ自動作成を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「縦棒横棒グラフ自動作成」を選択して実行する。

' ── 縦棒横棒グラフ自動作成 ──
Public Sub 縦棒横棒グラフ自動作成()
    Dim dataRange As Range
    Set dataRange = Application.InputBox("グラフ化するセル範囲を選択", "グラフ作成", Selection.Address, , , , , 8)
    Dim mode As String: mode = InputBox("種類: 1=縦棒、2=横棒", "グラフ作成", "1")
    Dim chartType As XlChartType
    If mode = "1" Then chartType = xlColumnClustered Else chartType = xlBarClustered
    ActiveSheet.Shapes.AddChart2(201, chartType).Chart.SetSourceData dataRange
End Sub




