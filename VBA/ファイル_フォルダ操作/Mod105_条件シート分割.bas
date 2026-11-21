Attribute VB_Name = "Mod105_条件シート分割"
Option Explicit
' ファイル処理概要: ファイル・フォルダを操作、ブック・シートを処理、辞書でキー照合を実行する。
' ファイルないコード処理詳細解説: 特定条件ごとにシート分割別ファイル出力を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、Windowsファイルシステム、必要な実行権限。
' ファイル実行例: ExcelのAlt+F8から「特定条件ごとにシート分割別ファイル出力」を選択して実行する。

' ── 特定条件ごとにシート分割別ファイル出力 ──
Public Sub 特定条件ごとにシート分割別ファイル出力()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim keyColumn As Variant
    keyColumn = Application.InputBox("区分に使う列番号を入力してください", "条件分割", 1, , , , , 1)
    If keyColumn = False Then Exit Sub
    If Not IsNumeric(keyColumn) Then Exit Sub

    Dim destinationFolder As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show <> -1 Then Exit Sub
        destinationFolder = .SelectedItems(1)
    End With

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim groups As Object: Set groups = CreateObject("Scripting.Dictionary")
    groups.CompareMode = vbTextCompare
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, CLng(keyColumn)).End(xlUp).Row
    Dim r As Long
    For r = 2 To lastRow
        Dim key As String: key = CStr(ws.Cells(r, CLng(keyColumn)).Value)
        If key <> "" And Not groups.Exists(key) Then groups.Add key, True
    Next r

    Dim item As Variant
    For Each item In groups.Keys
        Dim temp As Worksheet
        Set temp = ws.Parent.Worksheets.Add
        ws.Rows(1).Copy temp.Rows(1)
        Dim outRow As Long: outRow = 2
        For r = 2 To lastRow
            If CStr(ws.Cells(r, CLng(keyColumn)).Value) = CStr(item) Then
                ws.Rows(r).Copy temp.Rows(outRow)
                outRow = outRow + 1
            End If
        Next r
        Dim tempWb As Workbook: Set tempWb = Workbooks.Add(1)
        temp.Copy Before:=tempWb.Worksheets(1)
        tempWb.Worksheets(1).Delete
        tempWb.SaveAs fso.BuildPath(destinationFolder, Replace(CStr(item), "/", "_") & ".xlsx"), xlOpenXMLWorkbook
        tempWb.Close False
        Application.DisplayAlerts = False
        temp.Delete
        Application.DisplayAlerts = True
    Next item
    MsgBox "条件ごとのブックを別ファイルに出力しました。", vbInformation
End Sub




