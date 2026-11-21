Attribute VB_Name = "Mod35_取込結果読込"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: エラー取込結果ファイル読込集計を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「エラー取込結果ファイル読込集計」を選択して実行する。

' ── エラー取込結果ファイル読込集計 ──
Public Sub エラー取込結果ファイル読込集計()
    Dim filePath As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .Filters.Add "CSV", "*.csv"
        If .Show <> -1 Then Exit Sub
        filePath = .SelectedItems(1)
    End With
    Dim temp As Workbook: Set temp = Workbooks.Open(filePath)
    Dim source As Worksheet: Set source = temp.Worksheets(1)
    Dim output As Worksheet: Set output = GetOrCreateSheet
    output.Cells.Clear
    source.UsedRange.Copy output.Range("A1")
    temp.Close False
End Sub

' ── GetOrCreateSheet ──
Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("取込結果")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "取込結果"
    End If
End Function




