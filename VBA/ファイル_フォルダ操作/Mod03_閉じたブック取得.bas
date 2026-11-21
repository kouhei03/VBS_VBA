Attribute VB_Name = "Mod03_閉じたブック取得"
Option Explicit
' ファイル処理概要: ADOでDBへ接続、レコードセットをシートへ転記、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: 閉じたブックからデータ取得を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、ADO、対象DBドライバまたはOLE DBプロバイダー、DB接続権限。
' ファイル実行例: ExcelのAlt+F8から「閉じたブックからデータ取得」を選択して実行する。

' ACE OLE DB経由で、Excelで開かずに別ブックの表を読み込む。
' 実行環境に Microsoft Access Database Engine (ACE) が必要。
' ── 閉じたブックからデータ取得 ──
Public Sub 閉じたブックからデータ取得()
    Dim sourcePath As String
    sourcePath = SelectExcelFile()
    If sourcePath = "" Then Exit Sub

    Dim sql As String
    sql = InputBox("実行するSELECT文を入力してください。" & vbCrLf & _
                   "例: SELECT [顧客名], [売上] FROM [Sheet1$] WHERE [売上] >= 10000", _
                   "取得するデータの指定")
    If StrPtr(sql) = 0 Or Trim$(sql) = "" Then Exit Sub
    If UCase$(Left$(Trim$(sql), 6)) <> "SELECT" Then
        MsgBox "SELECT文を入力してください。", vbExclamation
        Exit Sub
    End If

    Dim outputAddress As String
    outputAddress = InputBox("取得結果を書き出すセルを入力してください。", _
                             "出力先", "A1")
    If StrPtr(outputAddress) = 0 Or Trim$(outputAddress) = "" Then Exit Sub

    On Error GoTo Failed

    Dim connection As Object
    Dim recordset As Object
    Dim connectionString As String
    Dim providerName As String
    Dim excelFormat As String
    Dim outputCell As Range
    Dim i As Long
    Dim errorText As String

    Set outputCell = ActiveSheet.Range(outputAddress)
    Set connection = CreateObject("ADODB.Connection")
    Set recordset = CreateObject("ADODB.Recordset")

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
    On Error Resume Next
    connection.Open connectionString
    If Err.Number <> 0 Then
        Err.Clear
        providerName = "Microsoft.ACE.OLEDB.12.0"
        connectionString = "Provider=" & providerName & ";" & _
                           "Data Source=" & sourcePath & ";" & _
                           "Mode=Read;Extended Properties=""" & excelFormat & ";HDR=YES;IMEX=1"";"
        connection.Open connectionString
    End If
    On Error GoTo Failed

    If connection.State = 0 Then
        Err.Raise vbObjectError + 1100, , _
                  "ACE OLE DB Provider が利用できません。Microsoft Access Database Engineをインストールしてください。"
    End If

    recordset.Open sql, connection, 0, 1, 1

    For i = 0 To recordset.Fields.Count - 1
        outputCell.Offset(0, i).Value = recordset.Fields(i).Name
    Next i

    If Not recordset.EOF Then
        outputCell.Offset(1, 0).CopyFromRecordset recordset
    End If

    recordset.Close
    connection.Close
    MsgBox "データを取得しました。", vbInformation
    Exit Sub

Failed:
    errorText = Err.Description
    On Error Resume Next
    If Not recordset Is Nothing Then
        If recordset.State <> 0 Then recordset.Close
    End If
    If Not connection Is Nothing Then
        If connection.State <> 0 Then connection.Close
    End If
    On Error GoTo 0
    MsgBox "データ取得に失敗しました: " & errorText & vbCrLf & _
           "SELECT文のシート名・列名と、ACE Providerのインストールを確認してください。", vbCritical
End Sub

' ── SelectExcelFile ──
Private Function SelectExcelFile() As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "データを取得するExcelブックを選択してください"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel ブック", "*.xlsx;*.xlsm;*.xlsb;*.xls"
        If .Show = -1 Then SelectExcelFile = .SelectedItems(1)
    End With
End Function




