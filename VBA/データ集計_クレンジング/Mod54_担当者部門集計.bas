Attribute VB_Name = "Mod54_担当者部門集計"
Option Explicit
' ファイル処理概要: 辞書でキー照合を実行する。
' ファイルないコード処理詳細解説: 担当者別集計表を作成を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「担当者別集計表を作成」を選択して実行する。

' ── 担当者別集計表を作成 ──
Public Sub 担当者別集計表を作成()
    Dim keyCol As Long: keyCol = CLng(InputBox("担当者列番号", "担当者別集計", "2"))
    Dim valueCol As Long: valueCol = CLng(InputBox("集計する数値列番号", "担当者別集計", "3"))
    Dim src As Worksheet: Set src = ActiveSheet
    Dim out As Worksheet: Set out = GetOrCreateSheet(src.Parent, "担当者別集計")
    out.Cells.Clear
    out.Cells(1, 1).Value = "担当者": out.Cells(1, 2).Value = "件数": out.Cells(1, 3).Value = "合計"
    Dim dic As Object: Set dic = CreateObject("Scripting.Dictionary")
    Dim r As Long
    For r = 2 To GetLastRow(src)
        Dim key As String: key = CStr(src.Cells(r, keyCol).Value)
        If Len(key) > 0 Then
            If Not dic.Exists(key) Then dic.Add key, 0
            dic(key) = dic(key) + CDbl(IIf(IsNumeric(src.Cells(r, valueCol).Value), src.Cells(r, valueCol).Value, 0))
        End If
    Next r
    r = 2
    Dim k As Variant
    For Each k In dic.Keys
        out.Cells(r, 1).Value = k
        out.Cells(r, 2).Value = Application.WorksheetFunction.CountIf(src.Columns(keyCol), k)
        out.Cells(r, 3).Value = dic(k)
        r = r + 1
    Next k
End Sub




