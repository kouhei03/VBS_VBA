Attribute VB_Name = "Mod21_文字列分割"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 区切り文字で分割を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「区切り文字で分割」を選択して実行する。

' ── 区切り文字で分割 ──
Public Sub 区切り文字で分割()
    Dim text As String
    text = InputBox("分割する文字列を入力してください", "文字列分割", "東京,大阪,福岡")
    If StrPtr(text) = 0 Then Exit Sub

    Dim parts() As String
    parts = Split(text, ",")
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        Debug.Print i & ": " & parts(i)
    Next i
End Sub




