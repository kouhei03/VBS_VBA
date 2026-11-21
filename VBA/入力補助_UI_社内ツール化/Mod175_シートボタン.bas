Attribute VB_Name = "Mod175_シートボタン"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: シート操作ボタン有効無効制御を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「シート操作ボタン有効無効制御」を選択して実行する。

' ── シート操作ボタン有効無効制御 ──
Public Sub シート操作ボタン有効無効制御()
    Dim enabled As Boolean
    enabled = (MsgBox("操作ボタンを有効にしますか？", vbYesNo + vbQuestion) = vbYes)
    Dim c As Shape
    For Each c In ActiveSheet.Shapes
        If c.Type = msoFormControl Or c.Type = msoOLEControlObject Then c.ControlFormat.Enabled = enabled
    Next c
End Sub




