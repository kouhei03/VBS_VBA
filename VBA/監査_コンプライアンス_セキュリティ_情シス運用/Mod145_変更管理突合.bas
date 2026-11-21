Attribute VB_Name = "Mod145_変更管理突合"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: 変更管理台帳実施結果突合を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「変更管理台帳実施結果突合」を選択して実行する。

' ── 変更管理台帳実施結果突合 ──
Public Sub 変更管理台帳実施結果突合()
    Dim planned As Worksheet: Set planned = ActiveSheet
    Dim actualName As String: actualName = InputBox("実施結果シート名", "変更管理突合")
    Dim actual As Worksheet
    On Error Resume Next
    Set actual = ThisWorkbook.Worksheets(actualName)
    On Error GoTo 0
    If actual Is Nothing Then Exit Sub
    Dim r As Long
    For r = 2 To planned.Cells(planned.Rows.Count, 1).End(xlUp).Row
        Dim found As Range
        Set found = actual.Columns(1).Find(CStr(planned.Cells(r, 1).Value), , xlValues, xlWhole)
        If found Is Nothing Then planned.Rows(r).Interior.Color = RGB(255, 199, 206)
    Next r
End Sub




