Attribute VB_Name = "Mod34_クレンジング検証"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: 連携前クレンジング検証レポートを中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「連携前クレンジング検証レポート」を選択して実行する。

' ── 連携前クレンジング検証レポート ──
Public Sub 連携前クレンジング検証レポート()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim report As Worksheet: Set report = GetOrCreateSheet
    report.Cells.Clear
    report.Range("A1:C1").Value = Array("セル", "問題", "件数")
    Dim r As Long: r = 2
    Dim cell As Range
    For Each cell In ws.UsedRange
        If VarType(cell.Value) = vbString And Len(CStr(cell.Value)) <> Len(Trim$(CStr(cell.Value))) Then
            report.Cells(r, 1).Value = cell.Address
            report.Cells(r, 2).Value = "前後空白"
            r = r + 1
        End If
    Next cell
    report.Columns("A:C").AutoFit
End Sub

' ── GetOrCreateSheet ──
Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("事前検証")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "事前検証"
    End If
End Function




