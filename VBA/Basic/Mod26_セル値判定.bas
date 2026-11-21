Attribute VB_Name = "Mod26_セル値判定"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: セルが空か判定を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「セルが空か判定」を選択して実行する。

' ── セルが空か判定 ──
Public Function セルが空か判定(ByVal target As Range) As Boolean
    セルが空か判定 = IsEmpty(target.Value) Or Trim$(CStr(target.Value)) = ""
End Function

' ── セルがNullか判定 ──
Public Function セルがNullか判定(ByVal target As Range) As Boolean
    セルがNullか判定 = IsNull(target.Value)
End Function

' ── セルがエラーか判定 ──
Public Function セルがエラーか判定(ByVal target As Range) As Boolean
    セルがエラーか判定 = IsError(target.Value)
End Function




