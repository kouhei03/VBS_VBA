Attribute VB_Name = "Mod111_定期業務チェック"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: 定期業務日次週次月次年次チェックリスト生成を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「定期業務日次週次月次年次チェックリスト生成」を選択して実行する。

' ── 定期業務日次週次月次年次チェックリスト生成 ──
Public Sub 定期業務日次週次月次年次チェックリスト生成()
    Dim cycle As String
    cycle = InputBox("周期を入力: 日次 / 週次 / 月次 / 年次", "定期業務")
    If cycle = "" Then Exit Sub
    Dim ws As Worksheet: Set ws = GetOrCreateSheet
    Dim r As Long: r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(r, 1).Value = cycle
    ws.Cells(r, 2).Value = Format$(Now, "yyyy/mm/dd hh:mm")
    ws.Cells(r, 3).Value = "未実施"
End Sub

' ── GetOrCreateSheet ──
Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("定期業務チェック")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "定期業務チェック"
    End If
End Function




