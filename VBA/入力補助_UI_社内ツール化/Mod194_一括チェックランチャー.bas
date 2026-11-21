Attribute VB_Name = "Mod194_一括チェックランチャー"
Option Explicit
' ファイル処理概要: ブック・シートを処理を実行する。
' ファイルないコード処理詳細解説: 複数ファイル一括チェック整形ランチャーを中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「複数ファイル一括チェック整形ランチャー」を選択して実行する。

' ── 複数ファイル一括チェック整形ランチャー ──
Public Sub 複数ファイル一括チェック整形ランチャー()
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = True
        .Filters.Add "Excel", "*.xlsm;*.xlsx;*.xls"
        If .Show <> -1 Then Exit Sub
        Dim i As Long
        For i = 1 To .SelectedItems.Count
            Dim wb As Workbook: Set wb = Workbooks.Open(.SelectedItems(i), UpdateLinks:=0, ReadOnly:=False)
            wb.ActiveSheet.UsedRange.Columns.AutoFit
            wb.Save
            wb.Close
        Next i
    End With
End Sub




