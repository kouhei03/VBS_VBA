Attribute VB_Name = "Mod32_シート保護"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: シート保護を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「シート保護」を選択して実行する。

' ── シート保護 ──
Public Sub シート保護(ByVal ws As Worksheet, ByVal password As String)
    ws.Protect Password:=password, DrawingObjects:=True, Contents:=True, _
               Scenarios:=True, AllowFiltering:=True
End Sub

' ── シート保護解除 ──
Public Sub シート保護解除(ByVal ws As Worksheet, ByVal password As String)
    ws.Unprotect Password:=password
End Sub




