Attribute VB_Name = "Mod64_クロス集計"
Option Explicit
' ファイル処理概要: 辞書でキー照合を実行する。
' ファイルないコード処理詳細解説: クロス集計表を作成を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「クロス集計表を作成」を選択して実行する。

' ── クロス集計表を作成 ──
Public Sub クロス集計表を作成()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim rowCol As Long: rowCol = CLng(InputBox("行項目の列番号", "クロス集計", "1"))
    Dim colCol As Long: colCol = CLng(InputBox("列項目の列番号", "クロス集計", "2"))
    Dim valueCol As Long: valueCol = CLng(InputBox("集計値の列番号", "クロス集計", "3"))
    Dim out As Worksheet: Set out = GetOrCreateSheet(ws.Parent, "クロス集計")
    out.Cells.Clear

    Dim rowsDic As Object: Set rowsDic = CreateObject("Scripting.Dictionary")
    Dim colsDic As Object: Set colsDic = CreateObject("Scripting.Dictionary")
    Dim sums As Object: Set sums = CreateObject("Scripting.Dictionary")
    Dim r As Long
    For r = 2 To GetLastRow(ws)
        Dim rowKey As String: rowKey = CStr(ws.Cells(r, rowCol).Value)
        Dim colKey As String: colKey = CStr(ws.Cells(r, colCol).Value)
        rowsDic(rowKey) = True: colsDic(colKey) = True
        Dim key As String: key = rowKey & vbNullChar & colKey
        If Not sums.Exists(key) Then sums.Add key, 0
        If IsNumeric(ws.Cells(r, valueCol).Value) Then sums(key) = sums(key) + CDbl(ws.Cells(r, valueCol).Value)
    Next r
    Dim c As Long: c = 2
    Dim colName As Variant
    For Each colName In colsDic.Keys
        out.Cells(1, c).Value = colName
        c = c + 1
    Next colName
    r = 2
    Dim rowName As Variant
    For Each rowName In rowsDic.Keys
        out.Cells(r, 1).Value = rowName
        c = 2
        For Each colName In colsDic.Keys
            key = CStr(rowName) & vbNullChar & CStr(colName)
            out.Cells(r, c).Value = IIf(sums.Exists(key), sums(key), 0)
            c = c + 1
        Next colName
        r = r + 1
    Next rowName
End Sub




