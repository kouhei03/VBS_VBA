Attribute VB_Name = "Mod111_処理ログ"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: ファイル処理ログテキストExcel出力を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「ファイル処理ログテキストExcel出力」を選択して実行する。

' ── ファイル処理ログテキストExcel出力 ──
Public Sub ファイル処理ログテキストExcel出力()
    Dim message As String
    message = InputBox("ログに記録する内容を入力", "処理ログ")
    If Trim$(message) = "" Then Exit Sub
    Dim mode As String
    mode = InputBox("出力形式を入力: 1=テキスト、2=Excelシート", "処理ログ", "1")
    If mode = "1" Then
        Dim logPath As String
        logPath = ThisWorkbook.Path & "\file_processing.log"
        Dim fileNumber As Integer: fileNumber = FreeFile
        Open logPath For Append As #fileNumber
        Print #fileNumber, Format$(Now, "yyyy/mm/dd hh:mm:ss") & vbTab & message
        Close #fileNumber
        MsgBox "ログへ追記しました。" & vbCrLf & logPath, vbInformation
    ElseIf mode = "2" Then
        Dim ws As Worksheet: Set ws = GetOrCreateSheet("処理ログ")
        Dim r As Long: r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
        ws.Cells(r, 1).Value = Now
        ws.Cells(r, 2).Value = message
        ws.Cells(r, 1).NumberFormat = "yyyy/mm/dd hh:mm:ss"
    End If
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




