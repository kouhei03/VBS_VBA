Attribute VB_Name = "Mod99_CSV書出し"
Option Explicit
' ファイル処理概要: ファイル・フォルダを操作、ブック・シートを処理を実行する。
' ファイルないコード処理詳細解説: 選択範囲CSV書出しを中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、Windowsファイルシステム、必要な実行権限。
' ファイル実行例: ExcelのAlt+F8から「選択範囲CSV書出し」を選択して実行する。

' ── 選択範囲CSV書出し ──
Public Sub 選択範囲CSV書出し()
    Dim destinationFolder As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show <> -1 Then Exit Sub
        destinationFolder = .SelectedItems(1)
    End With
    Dim fileName As String
    fileName = InputBox("CSVファイル名を入力", "CSV書出し", "export.csv")
    If Trim$(fileName) = "" Then Exit Sub
    If LCase$(Right$(fileName, 4)) <> ".csv" Then fileName = fileName & ".csv"

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim temp As Workbook
    Set temp = Workbooks.Add(1)
    Selection.Copy temp.Worksheets(1).Range("A1")
    temp.SaveAs fso.BuildPath(destinationFolder, fileName), xlCSV
    temp.Close False
    MsgBox "CSVを書き出しました。", vbInformation
End Sub




