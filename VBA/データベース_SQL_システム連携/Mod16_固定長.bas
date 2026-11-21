Attribute VB_Name = "Mod16_固定長"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: 固定長バイト指定パディング生成を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「固定長バイト指定パディング生成」を選択して実行する。

' ── 固定長バイト指定パディング生成 ──
Public Sub 固定長バイト指定パディング生成()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim length As Variant
    length = Application.InputBox("各フィールドの固定文字数", "固定長", 20, , , , , 1)
    If length = False Or Not IsNumeric(length) Then Exit Sub
    Dim output As Worksheet: Set output = GetOrCreateSheet
    Dim r As Long
    For r = 1 To lastRow
        output.Cells(r, 1).Value = Left$(CStr(ws.Cells(r, 1).Value) & String$(CLng(length), " "), CLng(length))
    Next r
End Sub

' ── GetOrCreateSheet ──
Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("固定長出力")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "固定長出力"
    End If
End Function




