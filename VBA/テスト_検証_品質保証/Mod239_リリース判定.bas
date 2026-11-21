Attribute VB_Name = "Mod239_リリース判定"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: リリース判定品質チェックリスト集計を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「リリース判定品質チェックリスト集計」を選択して実行する。

' ── リリース判定品質チェックリスト集計 ──
Public Sub リリース判定品質チェックリスト集計()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim okCount As Long, ngCount As Long, checkCount As Long
    Dim r As Long
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        Select Case CStr(ws.Cells(r, 2).Value)
            Case "OK": okCount = okCount + 1
            Case "NG": ngCount = ngCount + 1
            Case "要確認": checkCount = checkCount + 1
        End Select
    Next r
    Dim out As Worksheet: Set out = GetOrCreateSheet
    out.Cells.Clear
    out.Range("A1:B1").Value = Array("判定", "件数")
    out.Cells(2, 1).Value = "OK": out.Cells(2, 2).Value = okCount
    out.Cells(3, 1).Value = "NG": out.Cells(3, 2).Value = ngCount
    out.Cells(4, 1).Value = "要確認": out.Cells(4, 2).Value = checkCount
End Sub

' ── GetOrCreateSheet ──
Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("リリース判定")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "リリース判定"
    End If
End Function




