Attribute VB_Name = "Mod18_改行コード"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 改行コード指定連携ファイル出力を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、ADO、対象DBドライバまたはOLE DBプロバイダー、DB接続権限。
' ファイル実行例: ExcelのAlt+F8から「改行コード指定連携ファイル出力」を選択して実行する。

' ── 改行コード指定連携ファイル出力 ──
Public Sub 改行コード指定連携ファイル出力()
    Dim mode As String: mode = InputBox("改行コード: 1=CRLF、2=LF", "改行コード", "1")
    If mode <> "1" And mode <> "2" Then Exit Sub
    Dim destination As String
    With Application.FileDialog(msoFileDialogSaveAs)
        If .Show <> -1 Then Exit Sub
        destination = .SelectedItems(1)
    End With
    Dim text As String: text = CStr(Selection.Value)
    If mode = "1" Then
        text = Replace(text, vbLf, vbCrLf)
    Else
        text = Replace(text, vbCrLf, vbLf)
    End If
    Dim stream As Object: Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText text
    stream.SaveToFile destination, 2
    stream.Close
End Sub




