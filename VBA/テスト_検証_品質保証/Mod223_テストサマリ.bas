Attribute VB_Name = "Mod223_テストサマリ"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: テスト実施率消化率不具合件数サマリを中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「テスト実施率消化率不具合件数サマリ」を選択して実行する。

' ── テスト実施率消化率不具合件数サマリ ──
Public Sub テスト実施率消化率不具合件数サマリ()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim total As Long: total = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row - 1
    Dim executed As Long, ngCount As Long
    Dim r As Long
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If CStr(ws.Cells(r, 4).Value) <> "" Then executed = executed + 1
        If CStr(ws.Cells(r, 4).Value) = "NG" Then ngCount = ngCount + 1
    Next r
    Dim out As Worksheet: Set out = GetOrCreateSheet
    out.Cells.Clear
    out.Range("A1:B1").Value = Array("指標", "値")
    out.Cells(2, 1).Value = "全件数"
    out.Cells(2, 2).Value = total
    out.Cells(3, 1).Value = "実施率"
    out.Cells(3, 2).Value = IIf(total > 0, executed / total, 0)
    out.Cells(4, 1).Value = "不具合件数"
    out.Cells(4, 2).Value = ngCount
    out.Cells(3, 2).NumberFormat = "0.0%"
End Sub

' ── GetOrCreateSheet ──
Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("テストサマリ")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "テストサマリ"
    End If
End Function




