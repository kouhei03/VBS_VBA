Attribute VB_Name = "Mod86_定型タスク展開"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: 週次月次タスクテンプレート展開を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「週次月次タスクテンプレート展開」を選択して実行する。

' ── 週次月次タスクテンプレート展開 ──
Public Sub 週次月次タスクテンプレート展開()
    Dim cycle As String
    cycle = InputBox("周期を入力: 週次 / 月次", "定型タスク")
    If cycle <> "週次" And cycle <> "月次" Then Exit Sub
    Dim ws As Worksheet: Set ws = GetOrCreateSheet
    Dim r As Long: r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(r, 1).Value = Format$(Now, "yyyy/mm/dd")
    ws.Cells(r, 2).Value = cycle & "定期業務"
    ws.Cells(r, 3).Value = "未着手"
End Sub

' ── GetOrCreateSheet ──
Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("定型タスク")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "定型タスク"
    End If
End Function




