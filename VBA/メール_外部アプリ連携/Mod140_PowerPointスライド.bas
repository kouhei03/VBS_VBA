Attribute VB_Name = "Mod140_PowerPointスライド"
Option Explicit
' ファイル処理概要: PowerPointを操作を実行する。
' ファイルないコード処理詳細解説: ExcelデータからPowerPointスライド自動生成を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、Microsoft PowerPoint、COMオートメーション利用可。
' ファイル実行例: ExcelのAlt+F8から「ExcelデータからPowerPointスライド自動生成」を選択して実行する。

' ── ExcelデータからPowerPointスライド自動生成 ──
Public Sub ExcelデータからPowerPointスライド自動生成()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim ppt As Object: Set ppt = CreateObject("PowerPoint.Application")
    ppt.Visible = True
    Dim presentation As Object: Set presentation = ppt.Presentations.Add
    Dim r As Long
    For r = 1 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        Dim slide As Object: Set slide = presentation.Slides.Add(presentation.Slides.Count + 1, 1)
        slide.Shapes(1).TextFrame.TextRange.Text = CStr(ws.Cells(r, 1).Value)
        If ws.Cells(r, 2).Value <> "" Then
            slide.Shapes(2).TextFrame.TextRange.Text = CStr(ws.Cells(r, 2).Value)
        End If
    Next r
    MsgBox "スライドを生成しました。", vbInformation
End Sub




