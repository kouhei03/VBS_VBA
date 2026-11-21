Attribute VB_Name = "Mod191_対象行ハイライト"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 対象行自動ハイライト完了行切替を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「対象行自動ハイライト完了行切替」を選択して実行する。

' ── 対象行自動ハイライト完了行切替 ──
Public Sub 対象行自動ハイライト完了行切替()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim r As Long
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If CStr(ws.Cells(r, 3).Value) = "完了" Then
            ws.Rows(r).Font.Strikethrough = True
            ws.Rows(r).Font.Color = RGB(128, 128, 128)
        Else
            ws.Rows(r).Interior.Color = RGB(255, 235, 156)
        End If
    Next r
End Sub




