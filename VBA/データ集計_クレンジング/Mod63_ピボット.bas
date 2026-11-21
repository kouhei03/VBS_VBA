Attribute VB_Name = "Mod63_ピボット"
Option Explicit
' ファイル処理概要: ブック・シートを処理を実行する。
' ファイルないコード処理詳細解説: ピボットテーブル更新を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「ピボットテーブル更新」を選択して実行する。

' ── ピボットテーブル更新 ──
Public Sub ピボットテーブル更新()
    Dim pt As PivotTable
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        For Each pt In ws.PivotTables
            pt.RefreshTable
        Next pt
    Next ws
    MsgBox "ピボットテーブルを更新しました。", vbInformation
End Sub




