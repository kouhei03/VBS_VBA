Attribute VB_Name = "Mod19_取込チェック"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: 取込ファイル必須文字数桁数チェックを中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「取込ファイル必須文字数桁数チェック」を選択して実行する。

' ── 取込ファイル必須文字数桁数チェック ──
Public Sub 取込ファイル必須文字数桁数チェック()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim report As Worksheet: Set report = GetOrCreateSheet
    report.Cells.Clear
    report.Range("A1:C1").Value = Array("セル", "項目", "エラー")
    Dim r As Long: r = 2
    Dim cell As Range
    For Each cell In ws.UsedRange
        If cell.Column = 1 And Trim$(CStr(cell.Value)) = "" And cell.Row > 1 Then
            report.Cells(r, 1).Value = cell.Address
            report.Cells(r, 2).Value = "必須項目"
            report.Cells(r, 3).Value = "空欄"
            r = r + 1
        End If
    Next cell
    report.Columns("A:C").AutoFit
End Sub

' ── GetOrCreateSheet ──
Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("取込チェック")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "取込チェック"
    End If
End Function




