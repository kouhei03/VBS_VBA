Attribute VB_Name = "Mod197_帳票入力チェック"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: 帳票作成前入力チェックエラー一覧表示を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「帳票作成前入力チェックエラー一覧表示」を選択して実行する。

' ── 帳票作成前入力チェックエラー一覧表示 ──
Public Sub 帳票作成前入力チェックエラー一覧表示()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim report As Worksheet: Set report = GetOrCreateSheet("入力エラー一覧")
    report.Cells.Clear
    report.Range("A1:C1").Value = Array("シート", "セル", "エラー")
    Dim r As Long: r = 2
    Dim cell As Range
    For Each cell In ws.UsedRange
        If Trim$(CStr(cell.Value)) = "" And cell.Row > 1 Then
            report.Cells(r, 1).Value = ws.Name
            report.Cells(r, 2).Value = cell.Address
            report.Cells(r, 3).Value = "空欄"
            r = r + 1
        End If
    Next cell
    report.Columns("A:C").AutoFit
    MsgBox r - 2 & " 件の空欄を検出しました。", vbInformation
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




