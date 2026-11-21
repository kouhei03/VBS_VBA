Attribute VB_Name = "Mod215_数式エラー検出"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: 計算式数式エラー検出を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「計算式数式エラー検出」を選択して実行する。

' ── 計算式数式エラー検出 ──
Public Sub 計算式数式エラー検出()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim output As Worksheet: Set output = GetOrCreateSheet
    output.Cells.Clear
    output.Range("A1:C1").Value = Array("セル", "エラー", "数式")
    Dim r As Long: r = 2
    Dim c As Range
    For Each c In ws.UsedRange
        If IsError(c.Value) Then
            output.Cells(r, 1).Value = c.Address
            output.Cells(r, 2).Value = CStr(c.Text)
            output.Cells(r, 3).Value = IIf(c.HasFormula, c.Formula, "")
            r = r + 1
        End If
    Next c
End Sub

' ── GetOrCreateSheet ──
Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("数式エラー")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "数式エラー"
    End If
End Function




