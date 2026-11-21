Attribute VB_Name = "Mod31_大口分割出力"
Option Explicit
' ファイル処理概要: ブック・シートを処理を実行する。
' ファイルないコード処理詳細解説: DB抽出大口データシートブック分割を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「DB抽出大口データシートブック分割」を選択して実行する。

' ── DB抽出大口データシートブック分割 ──
Public Sub DB抽出大口データシートブック分割()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim pageSize As Variant
    pageSize = Application.InputBox("1シートあたりの行数", "分割出力", 50000, , , , , 1)
    If pageSize = False Or Not IsNumeric(pageSize) Then Exit Sub
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim startRow As Long: startRow = 2
    Dim part As Long: part = 1
    Do While startRow <= lastRow
        Dim output As Worksheet: Set output = ThisWorkbook.Worksheets.Add
        output.Name = "分割_" & Format$(part, "000")
        ws.Rows(1).Copy output.Rows(1)
        ws.Range(startRow & ":" & WorksheetFunction.Min(startRow + CLng(pageSize) - 1, lastRow)).Copy output.Rows(2)
        startRow = startRow + CLng(pageSize)
        part = part + 1
    Loop
End Sub




