Attribute VB_Name = "Mod03_閉じたブック取得"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod03_閉じたブック取得
' - ファイル名: Mod03_閉じたブック取得.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' ACE OLE DB経由で、Excelで開かずに別ブックの表を読み込む。
' 実行環境に Microsoft Access Database Engine (ACE) が必要。
' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub 閉じたブックからデータ取得() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub 閉じたブックからデータ取得()
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim sourcePath As String
    sourcePath = SelectExcelFile()
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If sourcePath = "" Then Exit Sub

    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim sql As String
    sql = InputBox("実行するSELECT文を入力してください。" & vbCrLf & _
                   "例: SELECT [顧客名], [売上] FROM [Sheet1$] WHERE [売上] >= 10000", _
                   "取得するデータの指定")
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If StrPtr(sql) = 0 Or Trim$(sql) = "" Then Exit Sub
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If UCase$(Left$(Trim$(sql), 6)) <> "SELECT" Then
        MsgBox "SELECT文を入力してください。", vbExclamation
        Exit Sub
    End If

    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim outputAddress As String
    outputAddress = InputBox("取得結果を書き出すセルを入力してください。", _
                             "出力先", "A1")
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If StrPtr(outputAddress) = 0 Or Trim$(outputAddress) = "" Then Exit Sub

    ' [ブロック解説: エラー処理設定]
    ' エラー発生時の処理分岐を定義します。エラー番号、エラー内容、復旧処理、画面更新の復帰を確認してください。
    On Error GoTo Failed

    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim connection As Object
    Dim recordset As Object
    Dim connectionString As String
    Dim providerName As String
    Dim excelFormat As String
    Dim outputCell As Range
    Dim i As Long
    Dim errorText As String

    ' [ブロック解説: オブジェクト参照の設定]
    ' Workbook、Worksheet、Range、外部COMオブジェクトなどの参照を初期化します。オブジェクトが存在し、必要な参照ライブラリまたはアプリケーションが利用可能か確認します。
    Set outputCell = ActiveSheet.Range(outputAddress)
    ' [ブロック解説: オブジェクト参照の設定]
    ' Workbook、Worksheet、Range、外部COMオブジェクトなどの参照を初期化します。オブジェクトが存在し、必要な参照ライブラリまたはアプリケーションが利用可能か確認します。
    Set connection = CreateObject("ADODB.Connection")
    ' [ブロック解説: オブジェクト参照の設定]
    ' Workbook、Worksheet、Range、外部COMオブジェクトなどの参照を初期化します。オブジェクトが存在し、必要な参照ライブラリまたはアプリケーションが利用可能か確認します。
    Set recordset = CreateObject("ADODB.Recordset")

    ' [ブロック解説: 分岐条件の設定]
    ' 入力値やステータスに応じた分岐を設定します。未定義値の既定動作、すべきでない処理へのフォールバックを確認してください。
    Select Case LCase$(Mid$(sourcePath, InStrRev(sourcePath, ".") + 1))
        Case "xls"
            excelFormat = "Excel 8.0"
        Case "xlsb"
            excelFormat = "Excel 12.0"
        Case Else
            excelFormat = "Excel 12.0 Xml"
    End Select

    providerName = "Microsoft.ACE.OLEDB.16.0"
    connectionString = "Provider=" & providerName & ";" & _
                       "Data Source=" & sourcePath & ";" & _
                       "Mode=Read;Extended Properties=""" & excelFormat & ";HDR=YES;IMEX=1"";"
    ' [ブロック解説: エラー処理設定]
    ' エラー発生時の処理分岐を定義します。エラー番号、エラー内容、復旧処理、画面更新の復帰を確認してください。
    On Error Resume Next
    connection.Open connectionString
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If Err.Number <> 0 Then
        Err.Clear
        providerName = "Microsoft.ACE.OLEDB.12.0"
        connectionString = "Provider=" & providerName & ";" & _
                           "Data Source=" & sourcePath & ";" & _
                           "Mode=Read;Extended Properties=""" & excelFormat & ";HDR=YES;IMEX=1"";"
        connection.Open connectionString
    End If
    ' [ブロック解説: エラー処理設定]
    ' エラー発生時の処理分岐を定義します。エラー番号、エラー内容、復旧処理、画面更新の復帰を確認してください。
    On Error GoTo Failed

    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If connection.State = 0 Then
        Err.Raise vbObjectError + 1100, , _
                  "ACE OLE DB Provider が利用できません。Microsoft Access Database Engineをインストールしてください。"
    End If

    recordset.Open sql, connection, 0, 1, 1

    ' [ブロック解説: 繰り返し処理]
    ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
    For i = 0 To recordset.Fields.Count - 1
        outputCell.Offset(0, i).Value = recordset.Fields(i).Name
    Next i

    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If Not recordset.EOF Then
        outputCell.Offset(1, 0).CopyFromRecordset recordset
    End If

    recordset.Close
    connection.Close
    MsgBox "データを取得しました。", vbInformation
    Exit Sub

Failed:
    errorText = Err.Description
    ' [ブロック解説: エラー処理設定]
    ' エラー発生時の処理分岐を定義します。エラー番号、エラー内容、復旧処理、画面更新の復帰を確認してください。
    On Error Resume Next
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If Not recordset Is Nothing Then
        ' [ブロック解説: 条件分岐]
        ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
        If recordset.State <> 0 Then recordset.Close
    End If
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If Not connection Is Nothing Then
        ' [ブロック解説: 条件分岐]
        ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
        If connection.State <> 0 Then connection.Close
    End If
    ' [ブロック解説: エラー処理設定]
    ' エラー発生時の処理分岐を定義します。エラー番号、エラー内容、復旧処理、画面更新の復帰を確認してください。
    On Error GoTo 0
    MsgBox "データ取得に失敗しました: " & errorText & vbCrLf & _
           "SELECT文のシート名・列名と、ACE Providerのインストールを確認してください。", vbCritical
End Sub

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Private Function SelectExcelFile() As String の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Private Function SelectExcelFile() As String
    ' [ブロック解説: 一括設定ブロック]
    ' Withブロック内のオブジェクトへ、セル形式、印刷設定、オプションなどをまとめて設定します。設定対象と上書き範囲に注意します。
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "データを取得するExcelブックを選択してください"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel ブック", "*.xlsx;*.xlsm;*.xlsb;*.xls"
        ' [ブロック解説: 条件分岐]
        ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
        If .Show = -1 Then SelectExcelFile = .SelectedItems(1)
    End With
End Function

