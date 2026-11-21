Attribute VB_Name = "Mod30_シート表示制御"
Option Explicit
' ファイル処理概要: ブック・シートを処理を実行する。
' ファイルないコード処理詳細解説: シート表示を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「シート表示」を選択して実行する。

' ── シート表示 ──
Public Sub シート表示(ByVal ws As Worksheet)
    ws.Visible = xlSheetVisible
End Sub

' ── シート非表示 ──
Public Sub シート非表示(ByVal ws As Worksheet)
    If wbVisibleSheetCount(ws.Parent) <= 1 Then
        Err.Raise vbObjectError + 3001, , "表示中の最後のシートは非表示にできません。"
    End If
    ws.Visible = xlSheetHidden
End Sub

' ── シートを非常に隠す ──
Public Sub シートを非常に隠す(ByVal ws As Worksheet)
    If wbVisibleSheetCount(ws.Parent) <= 1 Then
        Err.Raise vbObjectError + 3002, , "表示中の最後のシートをVeryHiddenにはできません。"
    End If
    ws.Visible = xlSheetVeryHidden
End Sub

' ── wbVisibleSheetCount ──
Private Function wbVisibleSheetCount(ByVal wb As Workbook) As Long
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.Visible = xlSheetVisible Then wbVisibleSheetCount = wbVisibleSheetCount + 1
    Next ws
End Function




