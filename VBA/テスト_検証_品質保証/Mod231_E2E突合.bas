Attribute VB_Name = "Mod231_E2E突合"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 画面入力DB登録出力ファイルE2E突合を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「画面入力DB登録出力ファイルE2E突合」を選択して実行する。

' ── 画面入力DB登録出力ファイルE2E突合 ──
Public Sub 画面入力DB登録出力ファイルE2E突合()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim r As Long
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        ws.Cells(r, 4).Value = IIf(CStr(ws.Cells(r, 2).Value) = CStr(ws.Cells(r, 3).Value), "一致", "不一致")
    Next r
End Sub




