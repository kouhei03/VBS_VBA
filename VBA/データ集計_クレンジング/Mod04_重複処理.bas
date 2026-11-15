Attribute VB_Name = "Mod04_重複処理"
Option Explicit

'==================================================
' Mod04_重複処理
' 顧客リスト・コード一覧などから重複行を検出し、
' ハイライト（色付け） または 削除 する
'
' 前提: アクティブシートは1行目がヘッダーの表形式
'==================================================

Private Const HIGHLIGHT_COLOR As Long = &H9AFFFF   ' 薄い黄色 (RGB 255,255,154相当)

'--------------------------------------------------
' メイン: 重複行をハイライト（2件目以降に色付け）
'--------------------------------------------------
Public Sub 重複ハイライト()
    Dim keyCols As Variant
    keyCols = AskKeyColumns()
    If IsEmpty(keyCols) Then Exit Sub

    SpeedUp True
    On Error GoTo Finally

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim lastRow As Long, lastCol As Long
    lastRow = GetLastRow(ws)
    lastCol = GetLastCol(ws)

    Dim dic As Object
    Set dic = CreateObject("Scripting.Dictionary")

    Dim r As Long
    Dim key As String
    Dim dupCount As Long
    For r = 2 To lastRow
        key = BuildKey(ws, r, keyCols)
        If dic.Exists(key) Then
            ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol)).Interior.Color = HIGHLIGHT_COLOR
            dupCount = dupCount + 1
        Else
            dic(key) = r
        End If
    Next r

Finally:
    SpeedUp False
    If Err.Number <> 0 Then
        MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    Else
        MsgBox "重複行: " & dupCount & " 行をハイライトしました。", vbInformation
    End If
End Sub

'--------------------------------------------------
' メイン: 重複行を削除（最初の1件を残す）
'--------------------------------------------------
Public Sub 重複削除()
    Dim keyCols As Variant
    keyCols = AskKeyColumns()
    If IsEmpty(keyCols) Then Exit Sub

    If MsgBox("重複行を削除します（最初の1件は残します）。" & vbCrLf & _
              "元に戻せないため、事前にバックアップを推奨します。" & vbCrLf & vbCrLf & _
              "実行しますか？", vbYesNo + vbExclamation) <> vbYes Then Exit Sub

    SpeedUp True
    On Error GoTo Finally

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim lastRow As Long
    lastRow = GetLastRow(ws)

    Dim dic As Object
    Set dic = CreateObject("Scripting.Dictionary")

    Dim delRange As Range
    Dim r As Long
    Dim key As String
    Dim dupCount As Long
    For r = 2 To lastRow
        key = BuildKey(ws, r, keyCols)
        If dic.Exists(key) Then
            If delRange Is Nothing Then
                Set delRange = ws.Rows(r)
            Else
                Set delRange = Union(delRange, ws.Rows(r))
            End If
            dupCount = dupCount + 1
        Else
            dic(key) = r
        End If
    Next r

    If Not delRange Is Nothing Then delRange.Delete

Finally:
    SpeedUp False
    If Err.Number <> 0 Then
        MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    Else
        MsgBox "重複行: " & dupCount & " 行を削除しました。", vbInformation
    End If
End Sub

'--------------------------------------------------
' ハイライトを解除
'--------------------------------------------------
Public Sub ハイライト解除()
    ActiveSheet.UsedRange.Interior.ColorIndex = xlNone
    MsgBox "ハイライトを解除しました。", vbInformation
End Sub

'--------------------------------------------------
' 重複判定キーの列番号をユーザーに確認
' 複数列はカンマ区切り（例: 1,3）で複合キー判定
'--------------------------------------------------
Private Function AskKeyColumns() As Variant
    Dim ans As String
    ans = InputBox("重複判定に使う列番号を入力してください" & vbCrLf & _
                   "（例: A列のみ → 1 / A列+C列の複合キー → 1,3）", _
                   "キー列の指定", "1")
    If ans = "" Then
        AskKeyColumns = Empty
        Exit Function
    End If

    Dim parts As Variant
    parts = Split(ans, ",")

    Dim i As Long
    Dim cols() As Long
    ReDim cols(LBound(parts) To UBound(parts))
    For i = LBound(parts) To UBound(parts)
        If Not IsNumeric(Trim$(parts(i))) Then
            MsgBox "列番号の形式が不正です: " & ans, vbExclamation
            AskKeyColumns = Empty
            Exit Function
        End If
        cols(i) = CLng(Trim$(parts(i)))
    Next i
    AskKeyColumns = cols
End Function

'--------------------------------------------------
' 指定行の複合キー文字列を生成
' 表記揺れによる判定漏れを防ぐため CleanText を適用
'--------------------------------------------------
Private Function BuildKey(ByVal ws As Worksheet, ByVal r As Long, _
                          ByVal keyCols As Variant) As String
    Dim i As Long
    Dim key As String
    For i = LBound(keyCols) To UBound(keyCols)
        key = key & CleanText(CStr(ws.Cells(r, keyCols(i)).Value)) & vbNullChar
    Next i
    BuildKey = key
End Function
