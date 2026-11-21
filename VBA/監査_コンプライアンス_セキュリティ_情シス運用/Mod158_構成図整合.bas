Attribute VB_Name = "Mod158_構成図整合"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: システム構成図パラメータシート整合チェックを中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「システム構成図パラメータシート整合チェック」を選択して実行する。

' ── システム構成図パラメータシート整合チェック ──
Public Sub システム構成図パラメータシート整合チェック()
    Dim diagram As Worksheet: Set diagram = ActiveSheet
    Dim sheetName As String: sheetName = InputBox("パラメータシート名", "構成整合")
    Dim parameters As Worksheet
    On Error Resume Next
    Set parameters = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If parameters Is Nothing Then Exit Sub
    Dim r As Long
    For r = 2 To diagram.Cells(diagram.Rows.Count, 1).End(xlUp).Row
        Dim found As Range
        Set found = parameters.Columns(1).Find(CStr(diagram.Cells(r, 1).Value), , xlValues, xlWhole)
        If found Is Nothing Then diagram.Rows(r).Interior.Color = RGB(255, 199, 206)
    Next r
End Sub




