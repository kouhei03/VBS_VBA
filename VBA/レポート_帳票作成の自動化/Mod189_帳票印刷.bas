Attribute VB_Name = "Mod189_帳票印刷"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 帳票印刷プリンタ部数指定を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「帳票印刷プリンタ部数指定」を選択して実行する。

' ── 帳票印刷プリンタ部数指定 ──
Public Sub 帳票印刷プリンタ部数指定()
    Dim printerName As String: printerName = InputBox("プリンタ名（空欄で既定）", "帳票印刷")
    Dim copies As Variant
    copies = Application.InputBox("印刷部数", "帳票印刷", 1, , , , , 1)
    If copies = False Then Exit Sub
    If Trim$(printerName) <> "" Then ActiveSheet.Application.ActivePrinter = printerName
    ActiveSheet.PrintOut Copies:=CLng(copies), Collate:=True
End Sub




