Attribute VB_Name = "Mod173_条件付き書式"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 条件付き書式達成期限超過上下位設定を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「条件付き書式達成期限超過上下位設定」を選択して実行する。

' ── 条件付き書式達成期限超過上下位設定 ──
Public Sub 条件付き書式達成期限超過上下位設定()
    If Selection Is Nothing Then Exit Sub
    With Selection.FormatConditions
        .Delete
        .Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
        .Item(.Count).Interior.Color = RGB(255, 199, 206)
        .Add Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="100"
        .Item(.Count).Interior.Color = RGB(198, 239, 206)
        .AddTop10
        .Item(.Count).TopBottom.TopBottom = xlTop10Top
        .Item(.Count).Interior.Color = RGB(255, 235, 156)
    End With
    MsgBox "条件付き書式を設定しました。", vbInformation
End Sub




