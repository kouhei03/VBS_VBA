Attribute VB_Name = "Mod133_ベースライン差分"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: 現行設定ベースラインゴールデン差分を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「現行設定ベースラインゴールデン差分」を選択して実行する。

' ── 現行設定ベースラインゴールデン差分 ──
Public Sub 現行設定ベースラインゴールデン差分()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim standardName As String: standardName = InputBox("標準設定シート名", "ベースライン差分")
    Dim standard As Worksheet
    On Error Resume Next
    Set standard = ThisWorkbook.Worksheets(standardName)
    On Error GoTo 0
    If standard Is Nothing Then Exit Sub
    Dim r As Long
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        Dim found As Range
        Set found = standard.Columns(1).Find(CStr(ws.Cells(r, 1).Value), , xlValues, xlWhole)
        If found Is Nothing Or CStr(ws.Cells(r, 2).Value) <> CStr(found.Offset(0, 1).Value) Then
            ws.Rows(r).Interior.Color = RGB(255, 199, 206)
        End If
    Next r
End Sub




