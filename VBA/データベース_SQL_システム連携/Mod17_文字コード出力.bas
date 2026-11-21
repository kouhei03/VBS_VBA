Attribute VB_Name = "Mod17_文字コード出力"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 文字コード指定出力ファイル作成を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、ADO、対象DBドライバまたはOLE DBプロバイダー、DB接続権限。
' ファイル実行例: ExcelのAlt+F8から「文字コード指定出力ファイル作成」を選択して実行する。

' ── 文字コード指定出力ファイル作成 ──
Public Sub 文字コード指定出力ファイル作成()
    Dim destination As String
    With Application.FileDialog(msoFileDialogSaveAs)
        If .Show <> -1 Then Exit Sub
        destination = .SelectedItems(1)
    End With
    Dim charset As String: charset = InputBox("文字コード: 932=Shift-JIS、65001=UTF-8", "文字コード", "65001")
    Dim stream As Object: Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = IIf(charset = "932", "shift_jis", "utf-8")
    stream.Open
    stream.WriteText Selection.Value
    stream.SaveToFile destination, 2
    stream.Close
End Sub




