Attribute VB_Name = "Mod66_欠損値リスト"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 欠損値をリストアップを中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「欠損値をリストアップ」を選択して実行する。

' ── 欠損値をリストアップ ──
Public Sub 欠損値をリストアップ()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim out As Worksheet: Set out = GetOrCreateSheet(ws.Parent, "欠損値リスト")
    out.Cells.Clear
    out.Cells(1, 1).Value = "行": out.Cells(1, 2).Value = "列": out.Cells(1, 3).Value = "列名"
    Dim outputRow As Long: outputRow = 2
    Dim r As Long, c As Long
    For r = 2 To GetLastRow(ws)
        For c = 1 To GetLastCol(ws)
            If Len(Trim$(CStr(ws.Cells(r, c).Value))) = 0 Then
                out.Cells(outputRow, 1).Value = r
                out.Cells(outputRow, 2).Value = c
                out.Cells(outputRow, 3).Value = ws.Cells(1, c).Value
                outputRow = outputRow + 1
            End If
        Next c
    Next r
End Sub




