Attribute VB_Name = "Mod139_PowerPoint起動"
Option Explicit
' ファイル処理概要: PowerPointを操作、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: PowerPoint起動新規作成を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、Microsoft PowerPoint、COMオートメーション利用可。
' ファイル実行例: ExcelのAlt+F8から「PowerPoint起動新規作成」を選択して実行する。

' ── PowerPoint起動新規作成 ──
Public Sub PowerPoint起動新規作成()
    Dim ppt As Object
    On Error Resume Next
    Set ppt = GetObject(, "PowerPoint.Application")
    On Error GoTo 0
    If ppt Is Nothing Then Set ppt = CreateObject("PowerPoint.Application")
    ppt.Visible = True
    ppt.Presentations.Add
End Sub




