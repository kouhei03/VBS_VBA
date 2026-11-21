Attribute VB_Name = "Mod01_Outlook下書き"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod01_Outlook下書き
' - ファイル名: Mod01_Outlook下書き.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' [ブロック解説: 変数・定数・設定値の宣言]
' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
Private Const MAIL_SHEET As String = "メール送信"
Private Const DEADLINE_SHEET As String = "期限管理"

' メール送信シートの標準見出しを作成する。
' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub メール送信シートを準備() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub メール送信シートを準備()
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim ws As Worksheet
    ' [ブロック解説: オブジェクト参照の設定]
    ' Workbook、Worksheet、Range、外部COMオブジェクトなどの参照を初期化します。オブジェクトが存在し、必要な参照ライブラリまたはアプリケーションが利用可能か確認します。
    Set ws = GetOrCreateSheet(ThisWorkbook, MAIL_SHEET)
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If Application.WorksheetFunction.CountA(ws.Rows(1)) = 0 Then
        ws.Range("A1:F1").Value = HeaderArray( _
            "宛先(To)", "件名", "本文", "添付ファイルのフルパス", "作成対象(はい)", "結果")
    End If
    ws.Rows(1).Font.Bold = True
    ws.Columns("A:F").AutoFit
    MsgBox "「" & MAIL_SHEET & "」シートを準備しました。", vbInformation
End Sub

' 期限管理シートの標準見出しを作成する。
' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub 期限管理シートを準備() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub 期限管理シートを準備()
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim ws As Worksheet
    ' [ブロック解説: オブジェクト参照の設定]
    ' Workbook、Worksheet、Range、外部COMオブジェクトなどの参照を初期化します。オブジェクトが存在し、必要な参照ライブラリまたはアプリケーションが利用可能か確認します。
    Set ws = GetOrCreateSheet(ThisWorkbook, DEADLINE_SHEET)
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If Application.WorksheetFunction.CountA(ws.Rows(1)) = 0 Then
        ws.Range("A1:G1").Value = DeadlineHeaderArray()
    End If
    ws.Rows(1).Font.Bold = True
    ws.Columns("A:G").AutoFit
    MsgBox "「" & DEADLINE_SHEET & "」シートを準備しました。", vbInformation
End Sub

' 選択対象の行からメールを作成し、送信せずOutlookの下書きに保存する。
' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub メール下書きを一括作成() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub メール下書きを一括作成()
    ' [ブロック解説: エラー処理設定]
    ' エラー発生時の処理分岐を定義します。エラー番号、エラー内容、復旧処理、画面更新の復帰を確認してください。
    On Error GoTo Failed

    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim ws As Worksheet
    ' [ブロック解説: オブジェクト参照の設定]
    ' Workbook、Worksheet、Range、外部COMオブジェクトなどの参照を初期化します。オブジェクトが存在し、必要な参照ライブラリまたはアプリケーションが利用可能か確認します。
    Set ws = GetOrCreateSheet(ThisWorkbook, MAIL_SHEET)
    EnsureHeaders ws, Array("宛先(To)", "件名", "本文", "添付ファイルのフルパス", "作成対象(はい)", "結果")

    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim lastRow As Long
    lastRow = LastUsedRow(ws)
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If lastRow < 2 Then
        MsgBox "「" & MAIL_SHEET & "」シートにデータがありません。", vbInformation
        Exit Sub
    End If

    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim outlook As Object
    Dim mail As Object
    Dim rowNumber As Long
    Dim createdCount As Long, failedCount As Long, skippedCount As Long
    Dim rowError As String

    ' [ブロック解説: 繰り返し処理]
    ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
    For rowNumber = 2 To lastRow
        ' [ブロック解説: 条件分岐]
        ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
        If IsSelected(ws.Cells(rowNumber, 5).Value) Then
            ' [ブロック解説: 条件分岐]
            ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
            If Left$(CStr(ws.Cells(rowNumber, 6).Value), Len("下書き保存:")) = "下書き保存:" Then
                skippedCount = skippedCount + 1
            Else
                ' [ブロック解説: エラー処理設定]
                ' エラー発生時の処理分岐を定義します。エラー番号、エラー内容、復旧処理、画面更新の復帰を確認してください。
                On Error GoTo RowFailed

                ' [ブロック解説: 変数・定数・設定値の宣言]
                ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
                Dim recipient As String
                Dim attachmentPath As String
                recipient = Trim$(CStr(ws.Cells(rowNumber, 1).Value))
                attachmentPath = Trim$(CStr(ws.Cells(rowNumber, 4).Value))
                ' [ブロック解説: 条件分岐]
                ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
                If recipient = "" Then Err.Raise vbObjectError + 1200, , "宛先が空欄です。"
                ' [ブロック解説: 条件分岐]
                ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
                If attachmentPath <> "" And Dir$(attachmentPath, vbNormal Or vbHidden Or vbSystem) = "" Then
                    Err.Raise vbObjectError + 1201, , "添付ファイルが見つかりません: " & attachmentPath
                End If

                ' [ブロック解説: 条件分岐]
                ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
                If outlook Is Nothing Then Set outlook = GetOutlookApp()
                Set mail = outlook.CreateItem(0)
                ' [ブロック解説: 一括設定ブロック]
                ' Withブロック内のオブジェクトへ、セル形式、印刷設定、オプションなどをまとめて設定します。設定対象と上書き範囲に注意します。
                With mail
                    .To = recipient
                    .Subject = CStr(ws.Cells(rowNumber, 2).Value)
                    .Body = CStr(ws.Cells(rowNumber, 3).Value)
                    ' [ブロック解説: 条件分岐]
                    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
                    If attachmentPath <> "" Then .Attachments.Add attachmentPath
                    .Save
                End With

                ws.Cells(rowNumber, 6).Value = "下書き保存: " & Format$(Now, "yyyy/mm/dd hh:nn:ss")
                createdCount = createdCount + 1
            End If
        End If
NextRow:
        ' [ブロック解説: エラー処理設定]
        ' エラー発生時の処理分岐を定義します。エラー番号、エラー内容、復旧処理、画面更新の復帰を確認してください。
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
' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub 期限リマインド下書きを作成() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub 期限リマインド下書きを作成()
    ' [ブロック解説: エラー処理設定]
    ' エラー発生時の処理分岐を定義します。エラー番号、エラー内容、復旧処理、画面更新の復帰を確認してください。
    On Error GoTo Failed

    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim ws As Worksheet
    ' [ブロック解説: オブジェクト参照の設定]
    ' Workbook、Worksheet、Range、外部COMオブジェクトなどの参照を初期化します。オブジェクトが存在し、必要な参照ライブラリまたはアプリケーションが利用可能か確認します。
    Set ws = GetOrCreateSheet(ThisWorkbook, DEADLINE_SHEET)
    EnsureHeaders ws, Array("タスク・契約名", "期限", "担当者", "メールアドレス", _
                            "何日前から通知", "最終下書き日", "結果")

    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim lastRow As Long
    lastRow = LastUsedRow(ws)
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If lastRow < 2 Then
        MsgBox "「" & DEADLINE_SHEET & "」シートにデータがありません。", vbInformation
        Exit Sub
    End If

    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim outlook As Object
    Dim mail As Object
    Dim rowNumber As Long
    Dim createdCount As Long, failedCount As Long, skippedCount As Long
    Dim rowError As String

    ' [ブロック解説: 繰り返し処理]
    ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
    For rowNumber = 2 To lastRow
        ' [ブロック解説: 条件分岐]
        ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
        If Trim$(CStr(ws.Cells(rowNumber, 1).Value)) <> "" Then
            ' [ブロック解説: エラー処理設定]
            ' エラー発生時の処理分岐を定義します。エラー番号、エラー内容、復旧処理、画面更新の復帰を確認してください。
            On Error GoTo RowFailed

            ' [ブロック解説: 変数・定数・設定値の宣言]
            ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
            Dim dueDate As Date
            Dim alertDays As Long
            Dim recipient As String
            Dim taskName As String
            Dim ownerName As String

            ' [ブロック解説: 条件分岐]
            ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
            If Not IsDate(ws.Cells(rowNumber, 2).Value) Then _
                Err.Raise vbObjectError + 1210, , "期限が日付として認識できません。"
            ' [ブロック解説: 条件分岐]
            ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
            If Not IsNumeric(ws.Cells(rowNumber, 5).Value) Then _
                Err.Raise vbObjectError + 1211, , "「何日前から通知」に0以上の数値を入力してください。"
            alertDays = CLng(ws.Cells(rowNumber, 5).Value)
            ' [ブロック解説: 条件分岐]
            ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
            If alertDays < 0 Then Err.Raise vbObjectError + 1212, , "通知日数に負の数は指定できません。"

            dueDate = DateValue(CDate(ws.Cells(rowNumber, 2).Value))
            ' [ブロック解説: 条件分岐]
            ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
            If dueDate > Date + alertDays Then
                ws.Cells(rowNumber, 7).Value = "通知時期前"
                GoTo NextRow
            End If

            ' [ブロック解説: 条件分岐]
            ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
            If IsDate(ws.Cells(rowNumber, 6).Value) Then
                ' [ブロック解説: 条件分岐]
                ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
                If DateValue(CDate(ws.Cells(rowNumber, 6).Value)) = Date Then
                    ws.Cells(rowNumber, 7).Value = "本日作成済み"
                    skippedCount = skippedCount + 1
                    GoTo NextRow
                End If
            End If

            recipient = Trim$(CStr(ws.Cells(rowNumber, 4).Value))
            ' [ブロック解説: 条件分岐]
            ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
            If recipient = "" Then Err.Raise vbObjectError + 1213, , "メールアドレスが空欄です。"
            taskName = CStr(ws.Cells(rowNumber, 1).Value)
            ownerName = Trim$(CStr(ws.Cells(rowNumber, 3).Value))
            ' [ブロック解説: 条件分岐]
            ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
            If outlook Is Nothing Then Set outlook = GetOutlookApp()

            Set mail = outlook.CreateItem(0)
            ' [ブロック解説: 一括設定ブロック]
            ' Withブロック内のオブジェクトへ、セル形式、印刷設定、オプションなどをまとめて設定します。設定対象と上書き範囲に注意します。
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
        ' [ブロック解説: エラー処理設定]
        ' エラー発生時の処理分岐を定義します。エラー番号、エラー内容、復旧処理、画面更新の復帰を確認してください。
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

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Private Function BuildDeadlineBody(ByVal taskName As String, ByVal dueDate As Date, _ の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Private Function BuildDeadlineBody(ByVal taskName As String, ByVal dueDate As Date, _
                                   ByVal ownerName As String) As String
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim body As String
    body = "以下の期限が近づいています。" & vbCrLf & vbCrLf & _
           "対象: " & taskName & vbCrLf & _
           "期限: " & Format$(dueDate, "yyyy/mm/dd")
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If ownerName <> "" Then body = body & vbCrLf & "担当者: " & ownerName
    BuildDeadlineBody = body
End Function

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Private Function GetOutlookApp() As Object の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Private Function GetOutlookApp() As Object
    ' [ブロック解説: エラー処理設定]
    ' エラー発生時の処理分岐を定義します。エラー番号、エラー内容、復旧処理、画面更新の復帰を確認してください。
    On Error Resume Next
    ' [ブロック解説: オブジェクト参照の設定]
    ' Workbook、Worksheet、Range、外部COMオブジェクトなどの参照を初期化します。オブジェクトが存在し、必要な参照ライブラリまたはアプリケーションが利用可能か確認します。
    Set GetOutlookApp = GetObject(, "Outlook.Application")
    ' [ブロック解説: エラー処理設定]
    ' エラー発生時の処理分岐を定義します。エラー番号、エラー内容、復旧処理、画面更新の復帰を確認してください。
    On Error GoTo 0
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If GetOutlookApp Is Nothing Then
        ' [ブロック解説: オブジェクト参照の設定]
        ' Workbook、Worksheet、Range、外部COMオブジェクトなどの参照を初期化します。オブジェクトが存在し、必要な参照ライブラリまたはアプリケーションが利用可能か確認します。
        Set GetOutlookApp = CreateObject("Outlook.Application")
    End If
End Function

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Private Function IsSelected(ByVal value As Variant) As Boolean の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Private Function IsSelected(ByVal value As Variant) As Boolean
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim text As String
    text = LCase$(Trim$(CStr(value)))
    IsSelected = (text = "はい" Or text = "yes" Or text = "y" Or text = "1" Or text = "true")
End Function

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Private Function HeaderArray(ByVal a As String, ByVal b As String, ByVal c As String, _ の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Private Function HeaderArray(ByVal a As String, ByVal b As String, ByVal c As String, _
                             ByVal d As String, ByVal e As String, ByVal f As String) As Variant
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim result(1 To 1, 1 To 6) As Variant
    result(1, 1) = a
    result(1, 2) = b
    result(1, 3) = c
    result(1, 4) = d
    result(1, 5) = e
    result(1, 6) = f
    HeaderArray = result
End Function

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Private Function DeadlineHeaderArray() As Variant の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Private Function DeadlineHeaderArray() As Variant
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
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

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Private Sub EnsureHeaders(ByVal ws As Worksheet, ByVal expected As Variant) の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Private Sub EnsureHeaders(ByVal ws As Worksheet, ByVal expected As Variant)
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim i As Long
    ' [ブロック解説: 繰り返し処理]
    ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
    For i = LBound(expected) To UBound(expected)
        ' [ブロック解説: 条件分岐]
        ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
        If CStr(ws.Cells(1, i + 1).Value) <> CStr(expected(i)) Then
            Err.Raise vbObjectError + 1220, , _
                      "シート「" & ws.Name & "」の1行目が標準見出しと一致しません。READMEの列構成を確認してください。"
        End If
    Next i
End Sub

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Private Function LastUsedRow(ByVal ws As Worksheet) As Long の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Private Function LastUsedRow(ByVal ws As Worksheet) As Long
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim found As Range
    Set found = ws.Cells.Find(What:="*", LookIn:=xlFormulas, _
                              SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If Not found Is Nothing Then LastUsedRow = found.Row
End Function

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Private Function GetOrCreateSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Private Function GetOrCreateSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    ' [ブロック解説: エラー処理設定]
    ' エラー発生時の処理分岐を定義します。エラー番号、エラー内容、復旧処理、画面更新の復帰を確認してください。
    On Error Resume Next
    ' [ブロック解説: オブジェクト参照の設定]
    ' Workbook、Worksheet、Range、外部COMオブジェクトなどの参照を初期化します。オブジェクトが存在し、必要な参照ライブラリまたはアプリケーションが利用可能か確認します。
    Set GetOrCreateSheet = wb.Worksheets(sheetName)
    ' [ブロック解説: エラー処理設定]
    ' エラー発生時の処理分岐を定義します。エラー番号、エラー内容、復旧処理、画面更新の復帰を確認してください。
    On Error GoTo 0
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If GetOrCreateSheet Is Nothing Then
        ' [ブロック解説: オブジェクト参照の設定]
        ' Workbook、Worksheet、Range、外部COMオブジェクトなどの参照を初期化します。オブジェクトが存在し、必要な参照ライブラリまたはアプリケーションが利用可能か確認します。
        Set GetOrCreateSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        GetOrCreateSheet.Name = sheetName
    End If
End Function

