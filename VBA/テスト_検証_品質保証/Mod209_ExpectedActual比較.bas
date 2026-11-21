Attribute VB_Name = "Mod209_ExpectedActual比較"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: ExpectedActualセル単位一括比較を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「ExpectedActualセル単位一括比較」を選択して実行する。

' ── ExpectedActualセル単位一括比較 ──
Public Sub ExpectedActualセル単位一括比較()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim mismatch As Worksheet: Set mismatch = GetOrCreateSheet
    mismatch.Cells.Clear
    mismatch.Range("A1:D1").Value = Array("セル", "Expected", "Actual", "判定")
    Dim r As Long: r = 2
    Dim c As Range
    For Each c In Selection
        If c.Column = 1 Then
            Dim actualCell As Range: Set actualCell = c.Offset(0, 1)
            mismatch.Cells(r, 1).Value = c.Address
            mismatch.Cells(r, 2).Value = c.Value
            mismatch.Cells(r, 3).Value = actualCell.Value
            mismatch.Cells(r, 4).Value = IIf(CStr(c.Value) = CStr(actualCell.Value), "OK", "NG")
            r = r + 1
        End If
    Next c
End Sub

' ── GetOrCreateSheet ──
Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("比較結果")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "比較結果"
    End If
End Function




