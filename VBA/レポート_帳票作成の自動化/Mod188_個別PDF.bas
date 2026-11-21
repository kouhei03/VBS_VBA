Attribute VB_Name = "Mod188_個別PDF"
Option Explicit
' ファイル処理概要: PDF等へ固定形式出力、ブック・シートを処理、辞書でキー照合、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: 対象ごと個別PDF一括出力を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「対象ごと個別PDF一括出力」を選択して実行する。

' ── 対象ごと個別PDF一括出力 ──
Public Sub 対象ごと個別PDF一括出力()
    Dim folderPath As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show <> -1 Then Exit Sub
        folderPath = .SelectedItems(1)
    End With
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim keyColumn As Variant
    keyColumn = Application.InputBox("対象名の列番号", "個別PDF", 1, , , , , 1)
    If keyColumn = False Then Exit Sub
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, CLng(keyColumn)).End(xlUp).Row
    Dim groups As Object: Set groups = CreateObject("Scripting.Dictionary")
    Dim r As Long
    For r = 2 To lastRow
        groups(CStr(ws.Cells(r, CLng(keyColumn)).Value)) = True
    Next r
    Dim key As Variant
    For Each key In groups.Keys
        Dim temp As Worksheet: Set temp = GetOrCreateSheet("個別_" & Replace(CStr(key), "/", "_"))
        temp.Cells.Clear
        ws.Rows(1).Copy temp.Rows(1)
        Dim outRow As Long: outRow = 2
        For r = 2 To lastRow
            If CStr(ws.Cells(r, CLng(keyColumn)).Value) = CStr(key) Then
                ws.Rows(r).Copy temp.Rows(outRow)
                outRow = outRow + 1
            End If
        Next r
        temp.ExportAsFixedFormat xlTypePDF, folderPath & "\" & Replace(CStr(key), "/", "_") & ".pdf"
        Application.DisplayAlerts = False
        temp.Delete
        Application.DisplayAlerts = True
    Next key
    MsgBox "個別PDFを出力しました。", vbInformation
End Sub

' ── GetOrCreateSheet ──
Private Function GetOrCreateSheet(ByVal name As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    If Len(name) > 31 Then name = Left$(name, 31)
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = name
    End If
End Function




