Attribute VB_Name = "Mod132_構成情報比較"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: サーバー機器構成設定比較を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「サーバー機器構成設定比較」を選択して実行する。

' ── サーバー機器構成設定比較 ──
Public Sub サーバー機器構成設定比較()
    Dim current As Worksheet: Set current = ActiveSheet
    Dim baselineName As String: baselineName = InputBox("比較先シート名", "構成比較")
    Dim baseline As Worksheet
    On Error Resume Next
    Set baseline = ThisWorkbook.Worksheets(baselineName)
    On Error GoTo 0
    If baseline Is Nothing Then Exit Sub
    Dim r As Long
    For r = 2 To current.Cells(current.Rows.Count, 1).End(xlUp).Row
        Dim key As String: key = CStr(current.Cells(r, 1).Value)
        Dim found As Range
        Set found = baseline.Columns(1).Find(key, , xlValues, xlWhole)
        If found Is Nothing Or CStr(current.Cells(r, 3).Value) <> CStr(found.Offset(0, 2).Value) Then
            current.Rows(r).Interior.Color = RGB(255, 199, 206)
        End If
    Next r
End Sub




