Attribute VB_Name = "Mod141_グラフPPT貼付"
Option Explicit
' ファイル処理概要: PowerPointを操作を実行する。
' ファイルないコード処理詳細解説: ExcelグラフをPowerPointへ一括貼付を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、Microsoft PowerPoint、COMオートメーション利用可。
' ファイル実行例: ExcelのAlt+F8から「ExcelグラフをPowerPointへ一括貼付」を選択して実行する。

' ── ExcelグラフをPowerPointへ一括貼付 ──
Public Sub ExcelグラフをPowerPointへ一括貼付()
    Dim ppt As Object: Set ppt = CreateObject("PowerPoint.Application")
    ppt.Visible = True
    Dim presentation As Object: Set presentation = ppt.Presentations.Add
    Dim chart As ChartObject
    For Each chart In ActiveSheet.ChartObjects
        chart.Copy
        Dim slide As Object: Set slide = presentation.Slides.Add(presentation.Slides.Count + 1, 1)
        slide.Shapes.Paste
    Next chart
    MsgBox "グラフを貼り付けました。", vbInformation
End Sub




