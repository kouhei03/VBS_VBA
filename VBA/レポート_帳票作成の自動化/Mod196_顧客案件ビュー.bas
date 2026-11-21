Attribute VB_Name = "Mod196_顧客案件ビュー"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: 顧客台帳案件管理定型ビュー作成を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「顧客台帳案件管理定型ビュー作成」を選択して実行する。

' ── 顧客台帳案件管理定型ビュー作成 ──
Public Sub 顧客台帳案件管理定型ビュー作成()
    Dim source As Worksheet: Set source = ActiveSheet
    Dim output As Worksheet: Set output = GetOrCreateSheet("顧客案件ビュー")
    output.Cells.Clear
    source.UsedRange.Copy output.Range("A1")
    output.Rows(1).Font.Bold = True
    output.Rows(1).Interior.Color = RGB(47, 85, 151)
    output.Rows(1).Font.Color = RGB(255, 255, 255)
    output.Columns.AutoFit
    With output.Sort
        .SortFields.Clear
        .SortFields.Add Key:=output.Range("A2"), Order:=xlAscending
        .SetRange output.UsedRange
        .Header = xlYes
        .Apply
    End With
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




