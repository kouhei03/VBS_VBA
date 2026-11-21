Attribute VB_Name = "Mod222_テスト結果判定"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: テスト結果OKNG要確認判定集計を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「テスト結果OKNG要確認判定集計」を選択して実行する。

' ── テスト結果OKNG要確認判定集計 ──
Public Sub テスト結果OKNG要確認判定集計()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim r As Long
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        Dim expected As String: expected = CStr(ws.Cells(r, 2).Value)
        Dim actual As String: actual = CStr(ws.Cells(r, 3).Value)
        If expected = "" Or actual = "" Then
            ws.Cells(r, 4).Value = "要確認"
        ElseIf expected = actual Then
            ws.Cells(r, 4).Value = "OK"
        Else
            ws.Cells(r, 4).Value = "NG"
        End If
    Next r
End Sub




