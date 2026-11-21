Attribute VB_Name = "Mod214_ブック構造比較"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: Excelブック複数シート構造シート名列名比較を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「Excelブック複数シート構造シート名列名比較」を選択して実行する。

' ── Excelブック複数シート構造シート名列名比較 ──
Public Sub Excelブック複数シート構造シート名列名比較()
    Dim file1 As String, file2 As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .Filters.Add "Excel", "*.xlsx;*.xlsm;*.xls"
        If .Show <> -1 Then Exit Sub
        file1 = .SelectedItems(1)
    End With
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .Filters.Add "Excel", "*.xlsx;*.xlsm;*.xls"
        If .Show <> -1 Then Exit Sub
        file2 = .SelectedItems(1)
    End With
    Dim wb1 As Workbook: Set wb1 = Workbooks.Open(file1, ReadOnly:=True)
    Dim wb2 As Workbook: Set wb2 = Workbooks.Open(file2, ReadOnly:=True)
    Dim output As Worksheet: Set output = GetOrCreateSheet
    output.Cells.Clear
    Dim r As Long: r = 1
    Dim ws As Worksheet
    For Each ws In wb1.Worksheets
        output.Cells(r, 1).Value = ws.Name
        output.Cells(r, 2).Value = wb2.Worksheets.Count
        r = r + 1
    Next ws
    wb1.Close False
    wb2.Close False
End Sub

' ── GetOrCreateSheet ──
Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("ブック構造比較")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "ブック構造比較"
    End If
End Function




