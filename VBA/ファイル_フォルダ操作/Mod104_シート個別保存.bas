Attribute VB_Name = "Mod104_シート個別保存"
Option Explicit
' ファイル処理概要: ブック・シートを処理を実行する。
' ファイルないコード処理詳細解説: 複数シートを個別ブックへ分割保存を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「複数シートを個別ブックへ分割保存」を選択して実行する。

' ── 複数シートを個別ブックへ分割保存 ──
Public Sub 複数シートを個別ブックへ分割保存()
    Dim destinationFolder As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show <> -1 Then Exit Sub
        destinationFolder = .SelectedItems(1)
    End With
    Application.DisplayAlerts = False
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        Dim temp As Workbook
        Set temp = Workbooks.Add(1)
        ws.Copy Before:=temp.Worksheets(1)
        temp.Worksheets(1).Delete
        temp.SaveAs destinationFolder & "\" & SafeSheetName(ws.Name) & ".xlsx", xlOpenXMLWorkbook
        temp.Close False
    Next ws
    Application.DisplayAlerts = True
    MsgBox "シートを個別ブックに分割保存しました。", vbInformation
End Sub

' ── SafeSheetName ──
Private Function SafeSheetName(ByVal name As String) As String
    Dim invalid As Variant: invalid = Array("\", "/", ":", "*", "?", Chr$(34), "<", ">", "|")
    Dim i As Long
    For i = LBound(invalid) To UBound(invalid)
        name = Replace(name, invalid(i), "_")
    Next i
    SafeSheetName = name
End Function




