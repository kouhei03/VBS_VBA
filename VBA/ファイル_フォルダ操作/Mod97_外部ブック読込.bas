Attribute VB_Name = "Mod97_外部ブック読込"
Option Explicit
' ファイル処理概要: ブック・シートを処理を実行する。
' ファイルないコード処理詳細解説: 外部ブック読み取り専用オープンを中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「外部ブック読み取り専用オープン」を選択して実行する。

' ── 外部ブック読み取り専用オープン ──
Public Sub 外部ブック読み取り専用オープン()
    Dim filePath As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .Filters.Clear
        .Filters.Add "Excel", "*.xls;*.xlsx;*.xlsm;*.xlsb"
        If .Show <> -1 Then Exit Sub
        filePath = .SelectedItems(1)
    End With
    Workbooks.Open fileName:=filePath, UpdateLinks:=0, ReadOnly:=True
End Sub




