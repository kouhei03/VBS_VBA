Attribute VB_Name = "Mod75_ランダム抽出"
Option Explicit
' ファイル処理概要: 辞書でキー照合を実行する。
' ファイルないコード処理詳細解説: ランダムサンプリング抽出を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「ランダムサンプリング抽出」を選択して実行する。

' ── ランダムサンプリング抽出 ──
Public Sub ランダムサンプリング抽出()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim sampleSize As Long: sampleSize = CLng(InputBox("抽出件数", "ランダム抽出", "10"))
    Dim out As Worksheet: Set out = GetOrCreateSheet(ws.Parent, "ランダムサンプル")
    out.Cells.Clear
    Dim lastRow As Long: lastRow = GetLastRow(ws)
    Dim lastCol As Long: lastCol = GetLastCol(ws)
    out.Range(out.Cells(1, 1), out.Cells(1, lastCol)).Value = _
        ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol)).Value
    Dim dic As Object: Set dic = CreateObject("Scripting.Dictionary")
    Randomize
    Dim outputRow As Long: outputRow = 2
    Do While outputRow < sampleSize + 2 And dic.Count < lastRow - 1
        Dim r As Long: r = Int((lastRow - 1) * Rnd) + 2
        If Not dic.Exists(r) Then
            dic.Add r, True
            out.Rows(outputRow).Value = ws.Rows(r).Value
            outputRow = outputRow + 1
        End If
    Loop
End Sub




