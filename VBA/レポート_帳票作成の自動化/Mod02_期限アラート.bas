Attribute VB_Name = "Mod02_期限アラート"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod02_期限アラート
' - ファイル名: Mod02_期限アラート.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' 期限管理シートから通知対象を抽出し、「期限アラート」シートに表示する。
' 日数は各行の「何日前から通知」列に入力する。
' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub 期限アラートを表示() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub 期限アラートを表示()
    ' [ブロック解説: エラー処理設定]
    ' エラー発生時の処理分岐を定義します。エラー番号、エラー内容、復旧処理、画面更新の復帰を確認してください。
    On Error GoTo Failed

    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim wsSource As Worksheet
    ' [ブロック解説: オブジェクト参照の設定]
    ' Workbook、Worksheet、Range、外部COMオブジェクトなどの参照を初期化します。オブジェクトが存在し、必要な参照ライブラリまたはアプリケーションが利用可能か確認します。
    Set wsSource = GetOrCreateSheet(ThisWorkbook, "期限管理")
    EnsureHeaders wsSource

    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim wsReport As Worksheet
    ' [ブロック解説: オブジェクト参照の設定]
    ' Workbook、Worksheet、Range、外部COMオブジェクトなどの参照を初期化します。オブジェクトが存在し、必要な参照ライブラリまたはアプリケーションが利用可能か確認します。
    Set wsReport = GetOrCreateSheet(ThisWorkbook, "期限アラート")
    wsReport.Cells.Clear
    wsReport.Cells(1, 1).Value = "タスク・契約名"
    wsReport.Cells(1, 2).Value = "期限"
    wsReport.Cells(1, 3).Value = "担当者"
    wsReport.Cells(1, 4).Value = "残日数"
    wsReport.Cells(1, 5).Value = "通知開始日数"
    wsReport.Rows(1).Font.Bold = True

    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim lastRow As Long
    lastRow = LastUsedRow(wsSource)

    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim rowNumber As Long
    Dim outputRow As Long
    Dim invalidCount As Long
    Dim dueDate As Date
    Dim alertDays As Long
    outputRow = 2

    ' [ブロック解説: 繰り返し処理]
    ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
    For rowNumber = 2 To lastRow
        ' [ブロック解説: 条件分岐]
        ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
        If Trim$(CStr(wsSource.Cells(rowNumber, 1).Value)) <> "" Then
            ' [ブロック解説: 条件分岐]
            ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
            If IsDate(wsSource.Cells(rowNumber, 2).Value) And _
               IsNumeric(wsSource.Cells(rowNumber, 5).Value) Then
                dueDate = DateValue(CDate(wsSource.Cells(rowNumber, 2).Value))
                alertDays = CLng(wsSource.Cells(rowNumber, 5).Value)

                ' [ブロック解説: 条件分岐]
                ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
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

    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim alertCount As Long
    alertCount = outputRow - 2
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
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

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Private Function InvalidSummary(ByVal invalidCount As Long) As String の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Private Function InvalidSummary(ByVal invalidCount As Long) As String
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If invalidCount > 0 Then
        InvalidSummary = vbCrLf & "期限または通知日数が不正な行: " & invalidCount & " 件"
    End If
End Function

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Private Sub EnsureHeaders(ByVal ws As Worksheet) の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Private Sub EnsureHeaders(ByVal ws As Worksheet)
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim expected As Variant
    expected = Array("タスク・契約名", "期限", "担当者", "メールアドレス", _
                     "何日前から通知", "最終下書き日", "結果")
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim i As Long
    ' [ブロック解説: 繰り返し処理]
    ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
    For i = LBound(expected) To UBound(expected)
        ' [ブロック解説: 条件分岐]
        ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
        If CStr(ws.Cells(1, i + 1).Value) <> CStr(expected(i)) Then
            Err.Raise vbObjectError + 1230, , _
                      "「期限管理」シートの1行目が標準見出しと一致しません。READMEの列構成を確認してください。"
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

