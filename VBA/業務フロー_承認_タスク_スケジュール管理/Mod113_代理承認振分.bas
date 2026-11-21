Attribute VB_Name = "Mod113_代理承認振分"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 代理承認不在期間タスク振分補助を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「代理承認不在期間タスク振分補助」を選択して実行する。

' ── 代理承認不在期間タスク振分補助 ──
Public Sub 代理承認不在期間タスク振分補助()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim absent As String: absent = InputBox("不在担当者名", "代理振分")
    Dim substitute As String: substitute = InputBox("代理担当者名", "代理振分")
    Dim r As Long
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If CStr(ws.Cells(r, 2).Value) = absent Then
            ws.Cells(r, 2).Value = substitute
            ws.Cells(r, 5).Value = "代理振分"
        End If
    Next r
End Sub




