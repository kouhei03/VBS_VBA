Attribute VB_Name = "Mod79_式を値へ変換"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 計算式を値へ一括変換を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「計算式を値へ一括変換」を選択して実行する。

' ── 計算式を値へ一括変換 ──
Public Sub 計算式を値へ一括変換()
    If TypeName(Selection) <> "Range" Then Exit Sub
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
End Sub




