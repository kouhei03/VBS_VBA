Attribute VB_Name = "Mod205_大量負荷データ"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: 大量負荷テストダミーデータ生成を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「大量負荷テストダミーデータ生成」を選択して実行する。

' ── 大量負荷テストダミーデータ生成 ──
Public Sub 大量負荷テストダミーデータ生成()
    Dim count As Variant: count = Application.InputBox("生成行数（例: 100000）", "負荷テスト", 100000, , , , , 1)
    If count = False Or Not IsNumeric(count) Then Exit Sub
    If CLng(count) > 1000000 Then
        MsgBox "安全上、最大100万行に制限します。", vbExclamation
        count = 1000000
    End If
    Dim ws As Worksheet: Set ws = GetOrCreateSheet
    ws.Cells.Clear
    ws.Range("A1").Resize(CLng(count), 3).Value = CreateLoadArray(CLng(count))
    ws.Columns("A:C").AutoFit
End Sub

' ── CreateLoadArray ──
Private Function CreateLoadArray(ByVal count As Long) As Variant
    Dim result() As Variant
    ReDim result(1 To count, 1 To 3)
    Dim i As Long
    For i = 1 To count
        result(i, 1) = i
        result(i, 2) = "DATA-" & Format$(i, "000000")
        result(i, 3) = i Mod 100
    Next i
    CreateLoadArray = result
End Function

' ── GetOrCreateSheet ──
Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("負荷テスト")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "負荷テスト"
    End If
End Function




