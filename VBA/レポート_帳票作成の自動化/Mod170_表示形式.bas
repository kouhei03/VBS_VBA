Attribute VB_Name = "Mod170_表示形式"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 数値表示形式設定を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「数値表示形式設定」を選択して実行する。

' ── 数値表示形式設定 ──
Public Sub 数値表示形式設定()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim mode As String
    mode = InputBox("形式を入力: 通貨 / 桁区切り / パーセント / 日付", "表示形式")
    Dim formatText As String
    Select Case mode
        Case "通貨": formatText = "[$￥-ja-JP]#,##0"
        Case "桁区切り": formatText = "#,##0"
        Case "パーセント": formatText = "0.0%"
        Case "日付": formatText = "yyyy/mm/dd"
        Case Else: Exit Sub
    End Select
    If Not Selection Is Nothing Then Selection.NumberFormat = formatText
End Sub




