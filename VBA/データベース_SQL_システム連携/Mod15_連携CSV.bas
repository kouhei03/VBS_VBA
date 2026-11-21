Attribute VB_Name = "Mod15_連携CSV"
Option Explicit
' ファイル処理概要: ブック・シートを処理を実行する。
' ファイルないコード処理詳細解説: 基幹システム取込CSV一括生成を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「基幹システム取込CSV一括生成」を選択して実行する。

' ── 基幹システム取込CSV一括生成 ──
Public Sub 基幹システム取込CSV一括生成()
    Dim destination As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show <> -1 Then Exit Sub
        destination = .SelectedItems(1)
    End With
    Dim temp As Workbook: Set temp = Workbooks.Add(1)
    ActiveSheet.UsedRange.Copy temp.Worksheets(1).Range("A1")
    temp.SaveAs destination & "\import_" & Format$(Now, "yyyymmdd_hhnnss") & ".csv", xlCSV
    temp.Close False
    MsgBox "連携CSVを生成しました。", vbInformation
End Sub




