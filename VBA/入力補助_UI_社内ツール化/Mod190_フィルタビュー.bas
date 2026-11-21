Attribute VB_Name = "Mod190_フィルタビュー"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: フィルタ一発クリアよく使うビュー切替を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「フィルタ一発クリアよく使うビュー切替」を選択して実行する。

' ── フィルタ一発クリアよく使うビュー切替 ──
Public Sub フィルタ一発クリアよく使うビュー切替()
    Dim ws As Worksheet: Set ws = ActiveSheet
    If ws.FilterMode Then ws.ShowAllData
    Dim mode As String: mode = InputBox("ビューを選択: 1=クリア、2=未完了", "フィルタビュー")
    If mode = "2" And ws.AutoFilterMode Then
        ws.UsedRange.AutoFilter Field:=3, Criteria1:="<>完了"
    End If
End Sub




