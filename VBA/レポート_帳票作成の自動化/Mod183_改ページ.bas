Attribute VB_Name = "Mod183_改ページ"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 手動改ページ挿入調整を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「手動改ページ挿入調整」を選択して実行する。

' ── 手動改ページ挿入調整 ──
Public Sub 手動改ページ挿入調整()
    If Selection Is Nothing Then Exit Sub
    ActiveSheet.HPageBreaks.Add Before:=Selection.Rows(1)
    MsgBox "選択行の前に改ページを挿入しました。", vbInformation
End Sub




