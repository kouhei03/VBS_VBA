Attribute VB_Name = "Mod13_二次元配列"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 配列を範囲へ一括書き出しを中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「配列を範囲へ一括書き出し」を選択して実行する。

' ── 配列を範囲へ一括書き出し ──
Public Sub 配列を範囲へ一括書き出し()
    Dim values(1 To 3, 1 To 2) As Variant
    values(1, 1) = "商品"
    values(1, 2) = "数量"
    values(2, 1) = "りんご"
    values(2, 2) = 10
    values(3, 1) = "みかん"
    values(3, 2) = 20
    ActiveSheet.Range("A1").Resize(UBound(values, 1), UBound(values, 2)).Value = values
End Sub




