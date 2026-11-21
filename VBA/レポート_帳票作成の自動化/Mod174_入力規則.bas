Attribute VB_Name = "Mod174_入力規則"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 入力規則ドロップダウン数値範囲設定を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「入力規則ドロップダウン数値範囲設定」を選択して実行する。

' ── 入力規則ドロップダウン数値範囲設定 ──
Public Sub 入力規則ドロップダウン数値範囲設定()
    If Selection Is Nothing Then Exit Sub
    Dim mode As String
    mode = InputBox("設定を入力: 1=リスト、2=0以上の数値", "入力規則")
    Selection.Validation.Delete
    If mode = "1" Then
        Dim listText As String
        listText = InputBox("カンマ区切りリストを入力", "入力規則")
        If listText <> "" Then Selection.Validation.Add xlValidateList, , , listText
    ElseIf mode = "2" Then
        Selection.Validation.Add xlValidateDecimal, xlValidAlertStop, xlBetween, "0", "100000000"
    End If
End Sub




