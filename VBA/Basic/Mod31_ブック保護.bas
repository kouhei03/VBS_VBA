Attribute VB_Name = "Mod31_ブック保護"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: ブック保護を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「ブック保護」を選択して実行する。

' ── ブック保護 ──
Public Sub ブック保護(ByVal wb As Workbook, ByVal password As String)
    wb.Protect Password:=password, Structure:=True, Windows:=False
End Sub

' ── ブック保護解除 ──
Public Sub ブック保護解除(ByVal wb As Workbook, ByVal password As String)
    wb.Unprotect Password:=password
End Sub




