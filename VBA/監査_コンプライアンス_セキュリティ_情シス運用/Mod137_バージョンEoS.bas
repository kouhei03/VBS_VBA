Attribute VB_Name = "Mod137_バージョンEoS"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: ソフトウェアOSミドルウェアバージョンEoS判定を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「ソフトウェアOSミドルウェアバージョンEoS判定」を選択して実行する。

' ── ソフトウェアOSミドルウェアバージョンEoS判定 ──
Public Sub ソフトウェアOSミドルウェアバージョンEoS判定()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim r As Long
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If IsDate(ws.Cells(r, 3).Value) And CDate(ws.Cells(r, 3).Value) < Date Then
            ws.Cells(r, 4).Value = "EoS"
            ws.Rows(r).Interior.Color = RGB(255, 199, 206)
        Else
            ws.Cells(r, 4).Value = "サポート中"
        End If
    Next r
End Sub




