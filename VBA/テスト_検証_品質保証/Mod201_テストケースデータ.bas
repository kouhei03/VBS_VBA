Attribute VB_Name = "Mod201_テストケースデータ"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: テストケース仕様テストデータ生成を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「テストケース仕様テストデータ生成」を選択して実行する。

' ── テストケース仕様テストデータ生成 ──
Public Sub テストケース仕様テストデータ生成()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim r As Long
    For r = 2 To lastRow
        ws.Cells(r, 3).Value = "入力: " & CStr(ws.Cells(r, 1).Value)
        ws.Cells(r, 4).Value = "期待: " & CStr(ws.Cells(r, 2).Value)
    Next r
End Sub




