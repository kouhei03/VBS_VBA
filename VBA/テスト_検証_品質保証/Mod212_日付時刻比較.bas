Attribute VB_Name = "Mod212_日付時刻比較"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 日付時刻フォーマットタイムゾーン比較検証を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「日付時刻フォーマットタイムゾーン比較検証」を選択して実行する。

' ── 日付時刻フォーマットタイムゾーン比較検証 ──
Public Sub 日付時刻フォーマットタイムゾーン比較検証()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim r As Long
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If IsDate(ws.Cells(r, 1).Value) And IsDate(ws.Cells(r, 2).Value) Then
            ws.Cells(r, 3).Value = IIf(DateDiff("s", CDate(ws.Cells(r, 1).Value), CDate(ws.Cells(r, 2).Value)) = 0, "OK", "NG")
        Else
            ws.Cells(r, 3).Value = "NG"
        End If
    Next r
End Sub




