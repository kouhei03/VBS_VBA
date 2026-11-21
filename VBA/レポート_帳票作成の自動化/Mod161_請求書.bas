Attribute VB_Name = "Mod161_請求書"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: 請求書自動作成明細転記を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「請求書自動作成明細転記」を選択して実行する。

' ── 請求書自動作成明細転記 ──
Public Sub 請求書自動作成明細転記()
    Dim source As Worksheet: Set source = ActiveSheet
    Dim template As Worksheet: Set template = ChooseSheet("請求書テンプレート")
    If template Is Nothing Then
        MsgBox "請求書テンプレートシートを選択できません。", vbExclamation
        Exit Sub
    End If
    Dim output As Worksheet: Set output = GetOrCreateSheet("請求書")
    template.UsedRange.Copy output.Range("A1")
    Dim lastRow As Long: lastRow = source.Cells(source.Rows.Count, 1).End(xlUp).Row
    source.Range(source.Cells(2, 1), source.Cells(lastRow, 5)).Copy output.Range("A10")
    output.Cells(5, 1).Value = source.Cells(1, 1).Value
    output.Columns("A:H").AutoFit
    MsgBox "請求書を作成しました。", vbInformation
End Sub

' ── ChooseSheet ──
Private Function ChooseSheet(ByVal name As String) As Worksheet
    On Error Resume Next
    Set ChooseSheet = ThisWorkbook.Worksheets(name)
End Function

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




