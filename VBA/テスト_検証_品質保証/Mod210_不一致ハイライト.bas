Attribute VB_Name = "Mod210_不一致ハイライト"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: 比較不一致ハイライト差異一覧生成を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「比較不一致ハイライト差異一覧生成」を選択して実行する。

' ── 比較不一致ハイライト差異一覧生成 ──
Public Sub 比較不一致ハイライト差異一覧生成()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim report As Worksheet: Set report = GetOrCreateSheet
    report.Cells.Clear
    report.Range("A1:D1").Value = Array("セル", "左値", "右値", "差異")
    Dim r As Long: r = 2
    Dim c As Range
    For Each c In ws.UsedRange
        If c.Column = 1 And VarType(c.Value) <> vbEmpty Then
            If CStr(c.Value) <> CStr(c.Offset(0, 1).Value) Then
                c.Interior.Color = RGB(255, 199, 206)
                c.Offset(0, 1).Interior.Color = RGB(255, 199, 206)
                report.Cells(r, 1).Value = c.Address
                report.Cells(r, 2).Value = c.Value
                report.Cells(r, 3).Value = c.Offset(0, 1).Value
                report.Cells(r, 4).Value = "NG"
                r = r + 1
            End If
        End If
    Next c
End Sub

' ── GetOrCreateSheet ──
Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("差異一覧")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "差異一覧"
    End If
End Function




