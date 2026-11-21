Attribute VB_Name = "Mod213_文字比較モード"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 空白改行大小文字比較モード切替を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「空白改行大小文字比較モード切替」を選択して実行する。

' ── 空白改行大小文字比較モード切替 ──
Public Sub 空白改行大小文字比較モード切替()
    Dim mode As String
    mode = InputBox("モード: 1=厳密、2=空白改行無視、3=大小文字無視", "文字比較")
    If Selection.Columns.Count < 2 Then Exit Sub
    Dim c As Range
    For Each c In Selection.Columns(1).Cells
        If c.Row > 0 And c.Row <= Selection.Rows.Count + Selection.Row - 1 Then
            Dim rightValue As String: rightValue = CStr(c.Offset(0, 1).Value)
            Dim leftValue As String: leftValue = CStr(c.Value)
            If mode = "2" Then
                leftValue = Replace(Replace(leftValue, " ", ""), vbCrLf, "")
                rightValue = Replace(Replace(rightValue, " ", ""), vbCrLf, "")
            ElseIf mode = "3" Then
                leftValue = LCase$(leftValue)
                rightValue = LCase$(rightValue)
            End If
            c.Offset(0, 2).Value = IIf(leftValue = rightValue, "OK", "NG")
        End If
    Next c
End Sub




