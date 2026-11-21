Attribute VB_Name = "Mod48_日付フォーマット"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 日付フォーマットを統一を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「日付フォーマットを統一」を選択して実行する。

' ── 日付フォーマットを統一 ──
Public Sub 日付フォーマットを統一()
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim rng As Range: Set rng = Selection
    Dim v As Variant
    For Each v In rng
        If IsDate(v.Value) Then
            v.Value = DateValue(v.Value)
            v.NumberFormat = "yyyy/mm/dd"
        End If
    Next v
End Sub




