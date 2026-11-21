Attribute VB_Name = "Mod56_シート突合"
Option Explicit
' ファイル処理概要: ブック・シートを処理、辞書でキー照合を実行する。
' ファイルないコード処理詳細解説: シート間キー項目を突合を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「シート間キー項目を突合」を選択して実行する。

' ── シート間キー項目を突合 ──
Public Sub シート間キー項目を突合()
    Dim leftName As String: leftName = InputBox("左側シート名", "キー突合")
    Dim rightName As String: rightName = InputBox("右側シート名", "キー突合")
    If leftName = "" Or rightName = "" Then Exit Sub
    If Not SheetExists(ThisWorkbook, leftName) Or Not SheetExists(ThisWorkbook, rightName) Then Exit Sub

    Dim leftWs As Worksheet: Set leftWs = ThisWorkbook.Worksheets(leftName)
    Dim rightWs As Worksheet: Set rightWs = ThisWorkbook.Worksheets(rightName)
    Dim dic As Object: Set dic = CreateObject("Scripting.Dictionary")
    Dim r As Long
    For r = 2 To GetLastRow(rightWs)
        dic(CStr(rightWs.Cells(r, 1).Value)) = True
    Next r
    Dim out As Worksheet: Set out = GetOrCreateSheet(ThisWorkbook, "突合結果")
    out.Cells.Clear
    out.Cells(1, 1).Value = "キー": out.Cells(1, 2).Value = "結果"
    Dim outputRow As Long: outputRow = 2
    For r = 2 To GetLastRow(leftWs)
        Dim key As String: key = CStr(leftWs.Cells(r, 1).Value)
        out.Cells(outputRow, 1).Value = key
        out.Cells(outputRow, 2).Value = IIf(dic.Exists(key), "一致", "不一致")
        outputRow = outputRow + 1
    Next r
End Sub




