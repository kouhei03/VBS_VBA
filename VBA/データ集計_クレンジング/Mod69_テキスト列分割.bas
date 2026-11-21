Attribute VB_Name = "Mod69_テキスト列分割"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: テキスト列を区切り文字で分割を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「テキスト列を区切り文字で分割」を選択して実行する。

' ── テキスト列を区切り文字で分割 ──
Public Sub テキスト列を区切り文字で分割()
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim separator As String: separator = InputBox("区切り文字", "テキスト分割", " ")
    If separator = "" Then Exit Sub
    Dim c As Range
    For Each c In Selection.Cells
        Dim parts As Variant: parts = Split(CStr(c.Value), separator)
        c.Resize(1, UBound(parts) + 1).Value = parts
    Next c
End Sub




