Attribute VB_Name = "Mod58_差分比較"
Option Explicit
' ファイル処理概要: ブック・シートを処理を実行する。
' ファイルないコード処理詳細解説: 新旧データの差分をハイライトを中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「新旧データの差分をハイライト」を選択して実行する。

' ── 新旧データの差分をハイライト ──
Public Sub 新旧データの差分をハイライト()
    Dim oldName As String: oldName = InputBox("旧データシート名", "差分比較")
    Dim newName As String: newName = InputBox("新データシート名", "差分比較")
    If Not SheetExists(ThisWorkbook, oldName) Or Not SheetExists(ThisWorkbook, newName) Then Exit Sub
    Dim oldWs As Worksheet: Set oldWs = ThisWorkbook.Worksheets(oldName)
    Dim newWs As Worksheet: Set newWs = ThisWorkbook.Worksheets(newName)
    Dim lastRow As Long: lastRow = Application.WorksheetFunction.Max(GetLastRow(oldWs), GetLastRow(newWs))
    Dim lastCol As Long: lastCol = Application.WorksheetFunction.Max(GetLastCol(oldWs), GetLastCol(newWs))
    Dim r As Long, c As Long
    For r = 1 To lastRow
        For c = 1 To lastCol
            If CStr(oldWs.Cells(r, c).Value) <> CStr(newWs.Cells(r, c).Value) Then
                newWs.Cells(r, c).Interior.Color = RGB(255, 230, 153)
            Else
                newWs.Cells(r, c).Interior.Pattern = xlNone
            End If
        Next c
    Next r
End Sub




