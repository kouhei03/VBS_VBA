Attribute VB_Name = "Mod02_クレンジング"
Option Explicit

'==================================================
' Mod02_クレンジング
' 表記揺れの一括修正
'   - 会社表記の統一（（株）→ 株式会社 など）
'   - 全角・半角の統一（英数字→半角、カナ→全角）
'   - 前後の不要なスペース削除、連続スペースの圧縮
'==================================================

'--------------------------------------------------
' メイン: 選択範囲（または使用範囲）を一括クレンジング
'--------------------------------------------------
Public Sub データクレンジング実行()
    Dim target As Range
    If TypeName(Selection) = "Range" And Selection.Cells.Count > 1 Then
        Set target = Selection
    Else
        Set target = ActiveSheet.UsedRange
    End If

    If MsgBox("対象: " & target.Address(False, False) & vbCrLf & _
              "・会社表記の統一" & vbCrLf & _
              "・英数字→半角 / カナ→全角" & vbCrLf & _
              "・前後スペース削除、連続スペース圧縮" & vbCrLf & vbCrLf & _
              "実行しますか？", vbYesNo + vbQuestion) <> vbYes Then Exit Sub

    SpeedUp True
    On Error GoTo Finally

    Dim arr As Variant
    arr = target.Value
    If Not IsArray(arr) Then
        arr = CleanText(CStr(arr))
        target.Value = arr
        GoTo Finally
    End If

    Dim r As Long, c As Long
    Dim changedCount As Long
    For r = LBound(arr, 1) To UBound(arr, 1)
        For c = LBound(arr, 2) To UBound(arr, 2)
            If VarType(arr(r, c)) = vbString Then
                Dim before As String, after As String
                before = arr(r, c)
                after = CleanText(before)
                If before <> after Then
                    arr(r, c) = after
                    changedCount = changedCount + 1
                End If
            End If
        Next c
    Next r
    target.Value = arr

Finally:
    SpeedUp False
    If Err.Number <> 0 Then
        MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    Else
        MsgBox "クレンジング完了: " & changedCount & " セルを修正しました。", vbInformation
    End If
End Sub

'--------------------------------------------------
' 1セル分のテキストをクレンジング
'--------------------------------------------------
Public Function CleanText(ByVal s As String) As String
    s = NormalizeWidth(s)
    s = NormalizeCompanyName(s)
    s = TrimSpaces(s)
    CleanText = s
End Function

'--------------------------------------------------
' 全角・半角の統一
'   英数字と記号 → 半角 / カタカナ → 全角
'--------------------------------------------------
Public Function NormalizeWidth(ByVal s As String) As String
    Dim result As String
    Dim i As Long
    Dim ch As String

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        Select Case True
            ' 全角英数字・全角スペース・主要な全角記号 → 半角
            Case ch Like "[０-９Ａ-Ｚａ-ｚ]" Or ch = "　" Or _
                 InStr("－．，＠＆（）／", ch) > 0
                result = result & StrConv(ch, vbNarrow)
            ' 半角カタカナ → 全角（濁点・半濁点は直前の文字と結合）
            Case ch Like "[ｦ-ﾟ]"
                If i < Len(s) And (Mid$(s, i + 1, 1) = "ﾞ" Or Mid$(s, i + 1, 1) = "ﾟ") Then
                    result = result & StrConv(ch & Mid$(s, i + 1, 1), vbWide)
                    i = i + 1
                Else
                    result = result & StrConv(ch, vbWide)
                End If
            Case Else
                result = result & ch
        End Select
    Next i
    NormalizeWidth = result
End Function

'--------------------------------------------------
' 会社表記の統一（すべて「株式会社」等の正式表記へ）
'--------------------------------------------------
Public Function NormalizeCompanyName(ByVal s As String) As String
    ' 半角カッコ表記（NormalizeWidth後は半角カッコになっている想定）
    s = Replace(s, "(株)", "株式会社")
    s = Replace(s, "(有)", "有限会社")
    s = Replace(s, "(合)", "合同会社")
    ' 全角カッコ表記（念のため両対応）
    s = Replace(s, "（株）", "株式会社")
    s = Replace(s, "（有）", "有限会社")
    s = Replace(s, "（合）", "合同会社")
    ' 略記
    s = Replace(s, "㈱", "株式会社")
    s = Replace(s, "㈲", "有限会社")
    NormalizeCompanyName = s
End Function

'--------------------------------------------------
' 前後スペース削除 + 文中の連続スペースを1つに圧縮
'--------------------------------------------------
Public Function TrimSpaces(ByVal s As String) As String
    s = Replace(s, "　", " ")           ' 全角スペース → 半角
    s = Replace(s, vbTab, " ")
    Do While InStr(s, "  ") > 0         ' 連続スペース圧縮
        s = Replace(s, "  ", " ")
    Loop
    TrimSpaces = Trim$(s)
End Function
