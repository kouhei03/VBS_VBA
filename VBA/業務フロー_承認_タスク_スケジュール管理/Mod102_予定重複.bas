Attribute VB_Name = "Mod102_予定重複"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: スケジュール予定重複検知を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「スケジュール予定重複検知」を選択して実行する。

' ── スケジュール予定重複検知 ──
Public Sub スケジュール予定重複検知()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim r As Long
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        Dim compareRow As Long
        For compareRow = r + 1 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            If CStr(ws.Cells(r, 1).Value) = CStr(ws.Cells(compareRow, 1).Value) _
               And CDate(ws.Cells(r, 2).Value) < CDate(ws.Cells(compareRow, 3).Value) _
               And CDate(ws.Cells(compareRow, 2).Value) < CDate(ws.Cells(r, 3).Value) Then
                ws.Rows(r).Interior.Color = RGB(255, 235, 156)
                ws.Rows(compareRow).Interior.Color = RGB(255, 235, 156)
            End If
        Next compareRow
    Next r
End Sub




