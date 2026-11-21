Attribute VB_Name = "Mod171_列幅行高"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 列幅行高自動調整余白確保を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「列幅行高自動調整余白確保」を選択して実行する。

' ── 列幅行高自動調整余白確保 ──
Public Sub 列幅行高自動調整余白確保()
    Dim ws As Worksheet: Set ws = ActiveSheet
    ws.UsedRange.Columns.AutoFit
    ws.UsedRange.Rows.AutoFit
    Dim c As Long
    For c = 1 To ws.UsedRange.Columns.Count
        If ws.Cells(1, c).EntireColumn.ColumnWidth > 30 Then
            ws.Cells(1, c).EntireColumn.ColumnWidth = 30
        End If
    Next c
    MsgBox "列幅と行高さを調整しました。", vbInformation
End Sub




