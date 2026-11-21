Attribute VB_Name = "Mod182_業務固有関数"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 業務コード整形を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「業務コード整形」を選択して実行する。

' ── 業務コード整形 ──
Public Function 業務コード整形(ByVal value As Variant) As String
    If VarType(value) = vbString Then
        業務コード整形 = UCase$(Replace(Trim$(CStr(value)), "-", ""))
    End If
End Function

' ── 営業日判定 ──
Public Function 営業日判定(ByVal value As Variant) As Boolean
    If IsDate(value) Then
        営業日判定 = (Weekday(CDate(value), vbMonday) <= 5)
    End If
End Function




