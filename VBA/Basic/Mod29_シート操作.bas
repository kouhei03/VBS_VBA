Attribute VB_Name = "Mod29_シート操作"
Option Explicit
' ファイル処理概要: ブック・シートを処理を実行する。
' ファイルないコード処理詳細解説: シート追加を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「シート追加」を選択して実行する。

' ── シート追加 ──
Public Function シート追加(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Set シート追加 = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    シート追加.Name = sheetName
End Function

' ── シート名称変更 ──
Public Sub シート名称変更(ByVal ws As Worksheet, ByVal newName As String)
    ws.Name = newName
End Sub

' ── シート削除 ──
Public Sub シート削除(ByVal ws As Worksheet)
    If wbSheetCount(ws.Parent) <= 1 Then
        Err.Raise vbObjectError + 2901, , "ブック内の最後のシートは削除できません。"
    End If
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
End Sub

' ── wbSheetCount ──
Private Function wbSheetCount(ByVal wb As Workbook) As Long
    wbSheetCount = wb.Worksheets.Count
End Function




