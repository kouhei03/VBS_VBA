Attribute VB_Name = "Mod129_営業時間外アクセス"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 営業時間外深夜休日アクセス抽出を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「営業時間外深夜休日アクセス抽出」を選択して実行する。

' ── 営業時間外深夜休日アクセス抽出 ──
Public Sub 営業時間外深夜休日アクセス抽出()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim r As Long
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If IsDate(ws.Cells(r, 2).Value) Then
            Dim d As Date: d = CDate(ws.Cells(r, 2).Value)
            If Hour(d) < 8 Or Hour(d) >= 19 Or Weekday(d, vbMonday) >= 6 Then
                ws.Rows(r).Interior.Color = RGB(255, 235, 156)
            End If
        End If
    Next r
End Sub




