Attribute VB_Name = "Mod220_回帰ベンチマーク"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: 回帰テストベンチマーク差異検出を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「回帰テストベンチマーク差異検出」を選択して実行する。

' ── 回帰テストベンチマーク差異検出 ──
Public Sub 回帰テストベンチマーク差異検出()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim benchmarkName As String: benchmarkName = InputBox("ベンチマークシート名", "回帰テスト")
    Dim benchmark As Worksheet
    On Error Resume Next
    Set benchmark = ThisWorkbook.Worksheets(benchmarkName)
    On Error GoTo 0
    If benchmark Is Nothing Then Exit Sub
    Dim r As Long
    For r = 1 To ws.UsedRange.Rows.Count
        If CStr(ws.Cells(r, 1).Value) <> CStr(benchmark.Cells(r, 1).Value) Then
            ws.Cells(r, 1).Interior.Color = RGB(255, 199, 206)
        End If
    Next r
End Sub




