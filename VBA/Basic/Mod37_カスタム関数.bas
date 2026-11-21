Attribute VB_Name = "Mod37_カスタム関数"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 税込金額を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「税込金額」を選択して実行する。

' 数式例: =税込金額(A1,0.1)
' ── 税込金額 ──
Public Function 税込金額(ByVal amount As Double, _
                         Optional ByVal taxRate As Double = 0.1) As Double
    税込金額 = amount * (1 + taxRate)
End Function

' 数式例: =文字列結合UDF(A1,B1," ")
' ── 文字列結合UDF ──
Public Function 文字列結合UDF(ByVal firstValue As Variant, _
                             ByVal secondValue As Variant, _
                             Optional ByVal separator As String = "") As String
    文字列結合UDF = CStr(firstValue) & separator & CStr(secondValue)
End Function




