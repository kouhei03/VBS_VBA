Attribute VB_Name = "Mod217_出力ファイル検証"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: CSV固定長レイアウト桁数終端検証を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、ADO、対象DBドライバまたはOLE DBプロバイダー、DB接続権限。
' ファイル実行例: ExcelのAlt+F8から「CSV固定長レイアウト桁数終端検証」を選択して実行する。

' ── CSV固定長レイアウト桁数終端検証 ──
Public Sub CSV固定長レイアウト桁数終端検証()
    Dim filePath As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .Filters.Add "Data", "*.csv;*.txt"
        If .Show <> -1 Then Exit Sub
        filePath = .SelectedItems(1)
    End With
    Dim stream As Object: Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile filePath
    Dim text As String: text = stream.ReadText
    stream.Close
    MsgBox "文字数: " & Len(text) & vbCrLf & "CRLF終端: " & IIf(Right$(text, 2) = vbCrLf, "あり", "なし"), vbInformation
End Sub




