Attribute VB_Name = "Mod206_組合せ展開"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: 組合せテストパラメータ展開を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「組合せテストパラメータ展開」を選択して実行する。

' ── 組合せテストパラメータ展開 ──
Public Sub 組合せテストパラメータ展開()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim out As Worksheet: Set out = GetOrCreateSheet
    out.Cells.Clear
    out.Range("A1:C1").Value = Array("組合せNo", "パラメータ1", "パラメータ2")
    Dim r As Long, outRow As Long: outRow = 2
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        Dim p As Variant
        For Each p In Split(CStr(ws.Cells(r, 2).Value), ",")
            out.Cells(outRow, 1).Value = outRow - 1
            out.Cells(outRow, 2).Value = ws.Cells(r, 1).Value
            out.Cells(outRow, 3).Value = Trim$(CStr(p))
            outRow = outRow + 1
        Next p
    Next r
End Sub

' ── GetOrCreateSheet ──
Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("組合せテスト")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "組合せテスト"
    End If
End Function




