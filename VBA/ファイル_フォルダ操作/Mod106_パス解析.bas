Attribute VB_Name = "Mod106_パス解析"
Option Explicit
' ファイル処理概要: ファイル・フォルダを操作、ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: パス文字列からファイル名拡張子親フォルダ抽出を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、Windowsファイルシステム、必要な実行権限。
' ファイル実行例: ExcelのAlt+F8から「パス文字列からファイル名拡張子親フォルダ抽出」を選択して実行する。

' ── パス文字列からファイル名拡張子親フォルダ抽出 ──
Public Sub パス文字列からファイル名拡張子親フォルダ抽出()
    Dim filePath As String
    filePath = InputBox("解析するパスを入力してください", "パス解析")
    If Trim$(filePath) = "" Then Exit Sub
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim ws As Worksheet: Set ws = GetOrCreateSheet("パス解析")
    ws.Cells.Clear
    ws.Cells(1, 1).Value = "項目"
    ws.Cells(1, 2).Value = "値"
    ws.Cells(2, 1).Value = "フルパス"
    ws.Cells(2, 2).Value = filePath
    ws.Cells(3, 1).Value = "ファイル名"
    ws.Cells(3, 2).Value = fso.GetFileName(filePath)
    ws.Cells(4, 1).Value = "ベース名"
    ws.Cells(4, 2).Value = fso.GetBaseName(filePath)
    ws.Cells(5, 1).Value = "拡張子"
    ws.Cells(5, 2).Value = fso.GetExtensionName(filePath)
    ws.Cells(6, 1).Value = "親フォルダ"
    ws.Cells(6, 2).Value = fso.GetParentFolderName(filePath)
    ws.Columns("A:B").AutoFit
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




