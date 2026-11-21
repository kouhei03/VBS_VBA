Attribute VB_Name = "Mod00_Outlook共通"
Option Explicit
' ファイル処理概要: Outlookを操作、ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: GetOutlookを中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、クラシックOutlook、Outlookプロファイル。
' ファイル実行例: ExcelのAlt+F8から「GetOutlook」を選択して実行する。

' ── GetOutlook ──
Public Function GetOutlook() As Object
    On Error Resume Next
    Set GetOutlook = GetObject(, "Outlook.Application")
    On Error GoTo 0
    If GetOutlook Is Nothing Then Set GetOutlook = CreateObject("Outlook.Application")
End Function

' ── GetOrCreateSheet ──
Public Function GetOrCreateSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = wb.Worksheets(sheetName)
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        GetOrCreateSheet.Name = sheetName
    End If
End Function