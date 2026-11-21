Attribute VB_Name = "Mod192_進捗ダッシュボード"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: 進捗ダッシュボードKPIカード達成バー更新を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「進捗ダッシュボードKPIカード達成バー更新」を選択して実行する。

' ── 進捗ダッシュボードKPIカード達成バー更新 ──
Public Sub 進捗ダッシュボードKPIカード達成バー更新()
    Dim ws As Worksheet: Set ws = GetOrCreateSheet("ダッシュボード")
    ws.Cells.Clear
    ws.Range("A1:C1").Merge
    ws.Cells(1, 1).Value = "進捗ダッシュボード"
    ws.Cells(1, 1).Font.Size = 16
    ws.Cells(3, 1).Value = "KPI"
    ws.Cells(3, 2).Value = "達成率"
    ws.Cells(3, 3).Value = "進捗バー"
    ws.Cells(4, 1).Value = "全体進捗"
    ws.Cells(4, 2).Value = 0.75
    ws.Cells(4, 2).NumberFormat = "0%"
    ws.Cells(4, 3).Formula = "=REPT(""■"",ROUND(B4*20,0))"
    ws.Columns("A:C").AutoFit
End Sub

' ── GetOrCreateSheet ──
Private Function GetOrCreateSheet(ByVal name As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = name
    End If
End Function




