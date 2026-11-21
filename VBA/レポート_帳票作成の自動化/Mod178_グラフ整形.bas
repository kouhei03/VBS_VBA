Attribute VB_Name = "Mod178_グラフ整形"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: グラフ位置サイズタイトル凡例自動整形を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「グラフ位置サイズタイトル凡例自動整形」を選択して実行する。

' ── グラフ位置サイズタイトル凡例自動整形 ──
Public Sub グラフ位置サイズタイトル凡例自動整形()
    Dim titleText As String: titleText = InputBox("グラフタイトル", "グラフ整形")
    Dim chartObject As ChartObject
    For Each chartObject In ActiveSheet.ChartObjects
        With chartObject
            .Left = 50
            .Top = 120
            .Width = 600
            .Height = 360
            .Chart.HasTitle = True
            .Chart.ChartTitle.Text = titleText
            .Chart.HasLegend = True
            .Chart.Legend.Position = xlLegendPositionBottom
        End With
    Next chartObject
End Sub




