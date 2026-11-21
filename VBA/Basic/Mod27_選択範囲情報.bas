Attribute VB_Name = "Mod27_選択範囲情報"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 選択範囲情報を表示を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「選択範囲情報を表示」を選択して実行する。

' ── 選択範囲情報を表示 ──
Public Sub 選択範囲情報を表示()
    Dim rng As Range
    Set rng = Application.Selection
    MsgBox "セル数: " & rng.Cells.Count & vbCrLf & _
           "行数: " & rng.Rows.Count & vbCrLf & _
           "列数: " & rng.Columns.Count, vbInformation
End Sub




