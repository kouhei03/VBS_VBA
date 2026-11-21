Attribute VB_Name = "Mod59_条件行抽出"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 条件に合う行を別シート抽出を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「条件に合う行を別シート抽出」を選択して実行する。

' ── 条件に合う行を別シート抽出 ──
Public Sub 条件に合う行を別シート抽出()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim col As Long: col = CLng(InputBox("条件列番号", "条件抽出", "1"))
    Dim criterion As String: criterion = InputBox("抽出する値", "条件抽出")
    Dim out As Worksheet: Set out = GetOrCreateSheet(ws.Parent, "条件抽出")
    out.Cells.Clear
    Dim lastCol As Long: lastCol = GetLastCol(ws)
    out.Range("A1").Resize(1, lastCol).Value = ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol)).Value
    Dim outputRow As Long: outputRow = 2
    Dim r As Long
    For r = 2 To GetLastRow(ws)
        If CStr(ws.Cells(r, col).Value) = criterion Then
            ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol)).Copy out.Cells(outputRow, 1)
            outputRow = outputRow + 1
        End If
    Next r
End Sub




