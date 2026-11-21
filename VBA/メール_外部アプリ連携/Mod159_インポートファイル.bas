Attribute VB_Name = "Mod159_インポートファイル"
Option Explicit
' ファイル処理概要: ブック・シートを処理を実行する。
' ファイルないコード処理詳細解説: 外部システム連携インポートファイル自動生成を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「外部システム連携インポートファイル自動生成」を選択して実行する。

' ── 外部システム連携インポートファイル自動生成 ──
Public Sub 外部システム連携インポートファイル自動生成()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim destinationFolder As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show <> -1 Then Exit Sub
        destinationFolder = .SelectedItems(1)
    End With
    Dim fileName As String: fileName = "import_" & Format$(Now, "yyyymmdd_hhnnss") & ".csv"
    Dim temp As Workbook: Set temp = Workbooks.Add(1)
    ws.UsedRange.Copy temp.Worksheets(1).Range("A1")
    temp.SaveAs destinationFolder & "\" & fileName, xlCSVUTF8
    temp.Close False
    MsgBox "インポートファイルを生成しました。" & vbCrLf & destinationFolder & "\" & fileName, vbInformation
End Sub




