Attribute VB_Name = "Mod114_複数CSV結合"
Option Explicit
' ファイル処理概要: ファイル・フォルダを操作、ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: 複数CSV一括結合シート出力を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、Windowsファイルシステム、必要な実行権限。
' ファイル実行例: ExcelのAlt+F8から「複数CSV一括結合シート出力」を選択して実行する。

' ── 複数CSV一括結合シート出力 ──
Public Sub 複数CSV一括結合シート出力()
    Dim folderPath As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show <> -1 Then Exit Sub
        folderPath = .SelectedItems(1)
    End With
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim ws As Worksheet: Set ws = GetOrCreateSheet("CSV結合")
    ws.Cells.Clear
    Dim outRow As Long: outRow = 1
    Dim file As Object
    For Each file In fso.GetFolder(folderPath).Files
        If LCase$(fso.GetExtensionName(file.Name)) = "csv" Then
            Dim csvWb As Workbook
            Set csvWb = Workbooks.Open(file.Path, UpdateLinks:=0, ReadOnly:=True)
            Dim source As Worksheet: Set source = csvWb.Worksheets(1)
            Dim lastRow As Long: lastRow = source.Cells(source.Rows.Count, 1).End(xlUp).Row
            Dim lastColumn As Long: lastColumn = source.Cells(1, source.Columns.Count).End(xlToLeft).Column
            If outRow > 1 Then
                source.Range(source.Cells(2, 1), source.Cells(lastRow, lastColumn)).Copy ws.Cells(outRow, 1)
            Else
                source.Range(source.Cells(1, 1), source.Cells(lastRow, lastColumn)).Copy ws.Cells(outRow, 1)
            End If
            outRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
            csvWb.Close False
        End If
    Next file
    ws.Columns.AutoFit
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




