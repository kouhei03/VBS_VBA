Attribute VB_Name = "Mod02_期限アラート"
Option Explicit

' 期限管理シートから通知対象を抽出し、「期限アラート」シートに表示する。
' 日数は各行の「何日前から通知」列に入力する。
Public Sub 期限アラートを表示()
    On Error GoTo Failed

    Dim wsSource As Worksheet
    Set wsSource = GetOrCreateSheet(ThisWorkbook, "期限管理")
    EnsureHeaders wsSource

    Dim wsReport As Worksheet
    Set wsReport = GetOrCreateSheet(ThisWorkbook, "期限アラート")
    wsReport.Cells.Clear
    wsReport.Cells(1, 1).Value = "タスク・契約名"
    wsReport.Cells(1, 2).Value = "期限"
    wsReport.Cells(1, 3).Value = "担当者"
    wsReport.Cells(1, 4).Value = "残日数"
    wsReport.Cells(1, 5).Value = "通知開始日数"
    wsReport.Rows(1).Font.Bold = True

    Dim lastRow As Long
    lastRow = LastUsedRow(wsSource)

    Dim rowNumber As Long
    Dim outputRow As Long
    Dim invalidCount As Long
    Dim dueDate As Date
    Dim alertDays As Long
    outputRow = 2

    For rowNumber = 2 To lastRow
        If Trim$(CStr(wsSource.Cells(rowNumber, 1).Value)) <> "" Then
            If IsDate(wsSource.Cells(rowNumber, 2).Value) And _
               IsNumeric(wsSource.Cells(rowNumber, 5).Value) Then
                dueDate = DateValue(CDate(wsSource.Cells(rowNumber, 2).Value))
                alertDays = CLng(wsSource.Cells(rowNumber, 5).Value)

                If alertDays >= 0 And dueDate <= Date + alertDays Then
                    wsReport.Cells(outputRow, 1).Value = wsSource.Cells(rowNumber, 1).Value
                    wsReport.Cells(outputRow, 2).Value = dueDate
                    wsReport.Cells(outputRow, 3).Value = wsSource.Cells(rowNumber, 3).Value
                    wsReport.Cells(outputRow, 4).Value = DateDiff("d", Date, dueDate)
                    wsReport.Cells(outputRow, 5).Value = alertDays
                    outputRow = outputRow + 1
                ElseIf alertDays < 0 Then
                    invalidCount = invalidCount + 1
                End If
            Else
                invalidCount = invalidCount + 1
            End If
        End If
    Next rowNumber

    wsReport.Columns(2).NumberFormat = "yyyy/mm/dd"
    wsReport.Columns("A:E").AutoFit

    Dim alertCount As Long
    alertCount = outputRow - 2
    If alertCount = 0 Then
        MsgBox "通知対象はありません。" & InvalidSummary(invalidCount), vbInformation
    Else
        wsReport.Activate
        MsgBox alertCount & " 件の通知対象を「期限アラート」シートに表示しました。" & _
               InvalidSummary(invalidCount), vbExclamation
    End If
    Exit Sub

Failed:
    MsgBox "期限アラートの作成に失敗しました: " & Err.Description, vbCritical
End Sub

Private Function InvalidSummary(ByVal invalidCount As Long) As String
    If invalidCount > 0 Then
        InvalidSummary = vbCrLf & "期限または通知日数が不正な行: " & invalidCount & " 件"
    End If
End Function

Private Sub EnsureHeaders(ByVal ws As Worksheet)
    Dim expected As Variant
    expected = Array("タスク・契約名", "期限", "担当者", "メールアドレス", _
                     "何日前から通知", "最終下書き日", "結果")
    Dim i As Long
    For i = LBound(expected) To UBound(expected)
        If CStr(ws.Cells(1, i + 1).Value) <> CStr(expected(i)) Then
            Err.Raise vbObjectError + 1230, , _
                      "「期限管理」シートの1行目が標準見出しと一致しません。READMEの列構成を確認してください。"
        End If
    Next i
End Sub

Private Function LastUsedRow(ByVal ws As Worksheet) As Long
    Dim found As Range
    Set found = ws.Cells.Find(What:="*", LookIn:=xlFormulas, _
                              SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If Not found Is Nothing Then LastUsedRow = found.Row
End Function

Private Function GetOrCreateSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = wb.Worksheets(sheetName)
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        GetOrCreateSheet.Name = sheetName
    End If
End Function
