Attribute VB_Name = "Mod35_UserFormリスト設定"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: コンボボックスを設定を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「コンボボックスを設定」を選択して実行する。

' ── コンボボックスを設定 ──
Public Sub コンボボックスを設定(ByVal combo As Object, ParamArray items() As Variant)
    Dim i As Long
    combo.Clear
    For i = LBound(items) To UBound(items)
        combo.AddItem CStr(items(i))
    Next i
    If combo.ListCount > 0 Then combo.ListIndex = 0
End Sub

' ── リストボックスを設定 ──
Public Sub リストボックスを設定(ByVal listBox As Object, ByVal values As Variant)
    Dim i As Long
    listBox.Clear
    For i = LBound(values, 1) To UBound(values, 1)
        listBox.AddItem CStr(values(i, 1))
    Next i
End Sub




