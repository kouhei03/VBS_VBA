Attribute VB_Name = "Mod96_複数ブック読込"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: 複数Excelブック順次オープンと取得を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「複数Excelブック順次オープンと取得」を選択して実行する。

' ── 複数Excelブック順次オープンと取得 ──
Public Sub 複数Excelブック順次オープンと取得()
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "Excel", "*.xls;*.xlsx;*.xlsm;*.xlsb"
        If .Show <> -1 Then Exit Sub

        Dim output As Worksheet: Set output = GetOrCreateSheet("複数ブック読込")
        output.Cells.Clear
        output.Cells(1, 1).Value = "ブック"
        output.Cells(1, 2).Value = "シート"
        output.Cells(1, 3).Value = "最終行"
        output.Cells(1, 4).Value = "最終列"

        Dim i As Long, r As Long: r = 2
        Dim source As Workbook
        For i = 1 To .SelectedItems.Count
            Set source = Workbooks.Open(.SelectedItems(i), UpdateLinks:=0, ReadOnly:=True)
            Dim ws As Worksheet
            For Each ws In source.Worksheets
                output.Cells(r, 1).Value = source.Name
                output.Cells(r, 2).Value = ws.Name
                output.Cells(r, 3).Value = ws.Cells.Find("*", , xlFormulas, , xlByRows, xlPrevious).Row
                output.Cells(r, 4).Value = ws.Cells.Find("*", , xlFormulas, , xlByColumns, xlPrevious).Column
                r = r + 1
            Next ws
            source.Close False
        Next i
        output.Columns("A:D").AutoFit
    End With
End Sub

' ── GetOrCreateSheet ──
Private Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = sheetName
    End If
End Function




