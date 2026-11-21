Attribute VB_Name = "Mod22_文字列結合"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 文字列を結合を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「文字列を結合」を選択して実行する。

' ── 文字列を結合 ──
Public Function 文字列を結合(ByVal separator As String, ParamArray values() As Variant) As String
    Dim result As String
    Dim i As Long
    For i = LBound(values) To UBound(values)
        If Len(CStr(values(i))) > 0 Then
            If Len(result) > 0 Then result = result & separator
            result = result & CStr(values(i))
        End If
    Next i
    文字列を結合 = result
End Function




