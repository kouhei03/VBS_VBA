Attribute VB_Name = "Mod165_リアルタイム検証"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 文字数桁数数値範囲日付リアルタイム検証を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「文字数桁数数値範囲日付リアルタイム検証」を選択して実行する。

' ── 文字数桁数数値範囲日付リアルタイム検証 ──
Public Sub 文字数桁数数値範囲日付リアルタイム検証()
    If Selection Is Nothing Then Exit Sub
    Dim mode As String: mode = InputBox("検証形式: 1=文字数、2=数値範囲、3=日付", "入力検証")
    Dim value As String: value = CStr(Selection.Value)
    Select Case mode
        Case "1"
            If Len(value) > 20 Then MsgBox "20文字以内で入力してください。", vbExclamation
        Case "2"
            If Not IsNumeric(value) Or CDbl(value) < 0 Or CDbl(value) > 100 Then _
                MsgBox "0以上100以下の数値を入力してください。", vbExclamation
        Case "3"
            If Not IsDate(value) Then MsgBox "有効な日付を入力してください。", vbExclamation
    End Select
End Sub




