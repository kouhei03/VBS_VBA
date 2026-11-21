Attribute VB_Name = "Mod01_Outlook下書き"
Option Explicit
' ファイル処理概要: Outlookを操作、ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: メール送信シートを準備を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、クラシックOutlook、Outlookプロファイル。
' ファイル実行例: ExcelのAlt+F8から「メール送信シートを準備」を選択して実行する。

Private Const MAIL_SHEET As String = "メール送信"
Private Const DEADLINE_SHEET As String = "期限管理"

' メール送信シートの標準見出しを作成する。
' ── メール送信シートを準備 ──
Public Sub メール送信シートを準備()
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(ThisWorkbook, MAIL_SHEET)
    If Application.WorksheetFunction.CountA(ws.Rows(1)) = 0 Then
        ws.Range("A1:F1").Value = HeaderArray( _
            "宛先(To)", "件名", "本文", "添付ファイルのフルパス", "作成対象(はい)", "結果")
    End If
    ws.Rows(1).Font.Bold = True
    ws.Columns("A:F").AutoFit
    MsgBox "「" & MAIL_SHEET & "」シートを準備しました。", vbInformation
End Sub

' 期限管理シートの標準見出しを作成する。
' ── 期限管理シートを準備 ──
Public Sub 期限管理シートを準備()
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(ThisWorkbook, DEADLINE_SHEET)
    If Application.WorksheetFunction.CountA(ws.Rows(1)) = 0 Then
        ws.Range("A1:G1").Value = DeadlineHeaderArray()
    End If
    ws.Rows(1).Font.Bold = True
    ws.Columns("A:G").AutoFit
    MsgBox "「" & DEADLINE_SHEET & "」シートを準備しました。", vbInformation
End Sub

' 選択対象の行からメールを作成し、送信せずOutlookの下書きに保存する。
' ── メール下書きを一括作成 ──
Public Sub メール下書きを一括作成()
    On Error GoTo Failed

    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(ThisWorkbook, MAIL_SHEET)
    EnsureHeaders ws, Array("宛先(To)", "件名", "本文", "添付ファイルのフルパス", "作成対象(はい)", "結果")

    Dim lastRow As Long
    lastRow = LastUsedRow(ws)
    If lastRow < 2 Then
        MsgBox "「" & MAIL_SHEET & "」シートにデータがありません。", vbInformation
        Exit Sub
    End If

    Dim outlook As Object
    Dim mail As Object
    Dim rowNumber As Long
    Dim createdCount As Long, failedCount As Long, skippedCount As Long
    Dim rowError As String

    For rowNumber = 2 To lastRow
        If IsSelected(ws.Cells(rowNumber, 5).Value) Then
            If Left$(CStr(ws.Cells(rowNumber, 6).Value), Len("下書き保存:")) = "下書き保存:" Then
                skippedCount = skippedCount + 1
            Else
                On Error GoTo RowFailed

                Dim recipient As String
                Dim attachmentPath As String
                recipient = Trim$(CStr(ws.Cells(rowNumber, 1).Value))
                attachmentPath = Trim$(CStr(ws.Cells(rowNumber, 4).Value))
                If recipient = "" Then Err.Raise vbObjectError + 1200, , "宛先が空欄です。"
                If attachmentPath <> "" And Dir$(attachmentPath, vbNormal Or vbHidden Or vbSystem) = "" Then
                    Err.Raise vbObjectError + 1201, , "添付ファイルが見つかりません: " & attachmentPath
                End If

                If outlook Is Nothing Then Set outlook = GetOutlookApp()
                Set mail = outlook.CreateItem(0)
                With mail
                    .To = recipient
                    .Subject = CStr(ws.Cells(rowNumber, 2).Value)
                    .Body = CStr(ws.Cells(rowNumber, 3).Value)
                    If attachmentPath <> "" Then .Attachments.Add attachmentPath
                    .Save
                End With

                ws.Cells(rowNumber, 6).Value = "下書き保存: " & Format$(Now, "yyyy/mm/dd hh:nn:ss")
                createdCount = createdCount + 1
            End If
        End If
NextRow:
        On Error GoTo Failed
    Next rowNumber

    MsgBox "下書きを " & createdCount & " 件作成しました。" & vbCrLf & _
           "失敗: " & failedCount & " 件 / 作成済みのためスキップ: " & skippedCount & " 件" & vbCrLf & _
           "送信はしていません。Outlookの下書きを確認してください。", vbInformation
    Exit Sub

RowFailed:
    rowError = Err.Description
    Err.Clear
    ws.Cells(rowNumber, 6).Value = "エラー: " & rowError
    failedCount = failedCount + 1
    Resume NextRow

Failed:
    MsgBox "メール下書きの作成に失敗しました: " & Err.Description, vbCritical
End Sub

' 期限管理シートの対象行から期限リマインドを下書き保存する。
' ── 期限リマインド下書きを作成 ──
Public Sub 期限リマインド下書きを作成()
    On Error GoTo Failed

    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(ThisWorkbook, DEADLINE_SHEET)
    EnsureHeaders ws, Array("タスク・契約名", "期限", "担当者", "メールアドレス", _
                            "何日前から通知", "最終下書き日", "結果")

    Dim lastRow As Long
    lastRow = LastUsedRow(ws)
    If lastRow < 2 Then
        MsgBox "「" & DEADLINE_SHEET & "」シートにデータがありません。", vbInformation
        Exit Sub
    End If

    Dim outlook As Object
    Dim mail As Object
    Dim rowNumber As Long
    Dim createdCount As Long, failedCount As Long, skippedCount As Long
    Dim rowError As String

    For rowNumber = 2 To lastRow
        If Trim$(CStr(ws.Cells(rowNumber, 1).Value)) <> "" Then
            On Error GoTo RowFailed

            Dim dueDate As Date
            Dim alertDays As Long
            Dim recipient As String
            Dim taskName As String
            Dim ownerName As String

            If Not IsDate(ws.Cells(rowNumber, 2).Value) Then _
                Err.Raise vbObjectError + 1210, , "期限が日付として認識できません。"
            If Not IsNumeric(ws.Cells(rowNumber, 5).Value) Then _
                Err.Raise vbObjectError + 1211, , "「何日前から通知」に0以上の数値を入力してください。"
            alertDays = CLng(ws.Cells(rowNumber, 5).Value)
            If alertDays < 0 Then Err.Raise vbObjectError + 1212, , "通知日数に負の数は指定できません。"

            dueDate = DateValue(CDate(ws.Cells(rowNumber, 2).Value))
            If dueDate > Date + alertDays Then
                ws.Cells(rowNumber, 7).Value = "通知時期前"
                GoTo NextRow
            End If

            If IsDate(ws.Cells(rowNumber, 6).Value) Then
                If DateValue(CDate(ws.Cells(rowNumber, 6).Value)) = Date Then
                    ws.Cells(rowNumber, 7).Value = "本日作成済み"
                    skippedCount = skippedCount + 1
                    GoTo NextRow
                End If
            End If

            recipient = Trim$(CStr(ws.Cells(rowNumber, 4).Value))
            If recipient = "" Then Err.Raise vbObjectError + 1213, , "メールアドレスが空欄です。"
            taskName = CStr(ws.Cells(rowNumber, 1).Value)
            ownerName = Trim$(CStr(ws.Cells(rowNumber, 3).Value))
            If outlook Is Nothing Then Set outlook = GetOutlookApp()

            Set mail = outlook.CreateItem(0)
            With mail
                .To = recipient
                .Subject = "【期限リマインド】" & taskName
                .Body = BuildDeadlineBody(taskName, dueDate, ownerName)
                .Save
            End With

            ws.Cells(rowNumber, 6).Value = Date
            ws.Cells(rowNumber, 6).NumberFormat = "yyyy/mm/dd"
            ws.Cells(rowNumber, 7).Value = "下書き保存: " & Format$(Now, "yyyy/mm/dd hh:nn:ss")
            createdCount = createdCount + 1
        End If
NextRow:
        On Error GoTo Failed
    Next rowNumber

    MsgBox "期限リマインドの下書きを " & createdCount & " 件作成しました。" & vbCrLf & _
           "失敗: " & failedCount & " 件 / 本日作成済み: " & skippedCount & " 件" & vbCrLf & _
           "送信はしていません。Outlookの下書きを確認してください。", vbInformation
    Exit Sub

RowFailed:
    rowError = Err.Description
    Err.Clear
    ws.Cells(rowNumber, 7).Value = "エラー: " & rowError
    failedCount = failedCount + 1
    Resume NextRow

Failed:
    MsgBox "期限リマインド下書きの作成に失敗しました: " & Err.Description, vbCritical
End Sub

' ── BuildDeadlineBody ──
Private Function BuildDeadlineBody(ByVal taskName As String, ByVal dueDate As Date, _
                                   ByVal ownerName As String) As String
    Dim body As String
    body = "以下の期限が近づいています。" & vbCrLf & vbCrLf & _
           "対象: " & taskName & vbCrLf & _
           "期限: " & Format$(dueDate, "yyyy/mm/dd")
    If ownerName <> "" Then body = body & vbCrLf & "担当者: " & ownerName
    BuildDeadlineBody = body
End Function

' ── GetOutlookApp ──
Private Function GetOutlookApp() As Object
    On Error Resume Next
    Set GetOutlookApp = GetObject(, "Outlook.Application")
    On Error GoTo 0
    If GetOutlookApp Is Nothing Then
        Set GetOutlookApp = CreateObject("Outlook.Application")
    End If
End Function

' ── IsSelected ──
Private Function IsSelected(ByVal value As Variant) As Boolean
    Dim text As String
    text = LCase$(Trim$(CStr(value)))
    IsSelected = (text = "はい" Or text = "yes" Or text = "y" Or text = "1" Or text = "true")
End Function

' ── HeaderArray ──
Private Function HeaderArray(ByVal a As String, ByVal b As String, ByVal c As String, _
                             ByVal d As String, ByVal e As String, ByVal f As String) As Variant
    Dim result(1 To 1, 1 To 6) As Variant
    result(1, 1) = a
    result(1, 2) = b
    result(1, 3) = c
    result(1, 4) = d
    result(1, 5) = e
    result(1, 6) = f
    HeaderArray = result
End Function

' ── DeadlineHeaderArray ──
Private Function DeadlineHeaderArray() As Variant
    Dim result(1 To 1, 1 To 7) As Variant
    result(1, 1) = "タスク・契約名"
    result(1, 2) = "期限"
    result(1, 3) = "担当者"
    result(1, 4) = "メールアドレス"
    result(1, 5) = "何日前から通知"
    result(1, 6) = "最終下書き日"
    result(1, 7) = "結果"
    DeadlineHeaderArray = result
End Function

' ── EnsureHeaders ──
Private Sub EnsureHeaders(ByVal ws As Worksheet, ByVal expected As Variant)
    Dim i As Long
    For i = LBound(expected) To UBound(expected)
        If CStr(ws.Cells(1, i + 1).Value) <> CStr(expected(i)) Then
            Err.Raise vbObjectError + 1220, , _
                      "シート「" & ws.Name & "」の1行目が標準見出しと一致しません。READMEの列構成を確認してください。"
        End If
    Next i
End Sub

' ── LastUsedRow ──
Private Function LastUsedRow(ByVal ws As Worksheet) As Long
    Dim found As Range
    Set found = ws.Cells.Find(What:="*", LookIn:=xlFormulas, _
                              SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If Not found Is Nothing Then LastUsedRow = found.Row
End Function

' ── GetOrCreateSheet ──
Private Function GetOrCreateSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = wb.Worksheets(sheetName)
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        GetOrCreateSheet.Name = sheetName
    End If
End Function




