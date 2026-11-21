Attribute VB_Name = "Mod238_環境比較"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: テスト環境検証A検証B結果比較を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「テスト環境検証A検証B結果比較」を選択して実行する。

' ── テスト環境検証A検証B結果比較 ──
Public Sub テスト環境検証A検証B結果比較()
    Dim envA As String: envA = InputBox("検証Aシート名", "環境比較")
    Dim envB As String: envB = InputBox("検証Bシート名", "環境比較")
    Dim wsA As Worksheet, wsB As Worksheet
    On Error Resume Next
    Set wsA = ThisWorkbook.Worksheets(envA)
    Set wsB = ThisWorkbook.Worksheets(envB)
    On Error GoTo 0
    If wsA Is Nothing Or wsB Is Nothing Then Exit Sub
    Dim out As Worksheet: Set out = GetOrCreateSheet
    out.Cells.Clear
    out.Range("A1:C1").Value = Array("セル", "検証A", "検証B")
    Dim r As Long
    For r = 1 To Application.WorksheetFunction.Max(wsA.UsedRange.Rows.Count, wsB.UsedRange.Rows.Count)
        out.Cells(r + 1, 1).Value = "R" & r
        out.Cells(r + 1, 2).Value = wsA.Cells(r, 1).Value
        out.Cells(r + 1, 3).Value = wsB.Cells(r, 1).Value
    Next r
End Sub

' ── GetOrCreateSheet ──
Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("環境比較")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "環境比較"
    End If
End Function




