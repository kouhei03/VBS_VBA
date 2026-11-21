Attribute VB_Name = "Mod03_条件振り分け"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod03_条件振り分け
' - ファイル名: Mod03_条件振り分け.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

'==================================================
' Mod03_条件振り分け
' キー列（支店名・担当者名など）の値ごとにデータを
' 抽出し、別シート / 別ブックに自動分割する
'
' 前提: アクティブシートは1行目がヘッダーの表形式
'==================================================

'--------------------------------------------------
' メイン: キー列の値ごとに「別シート」へ振り分け
'--------------------------------------------------
' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub 条件別シート振り分け() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub 条件別シート振り分け()
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim keyCol As Long
    keyCol = AskKeyColumn()
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If keyCol = 0 Then Exit Sub

    SpeedUp True
    ' [ブロック解説: エラー処理設定]
    ' エラー発生時の処理分岐を定義します。エラー番号、エラー内容、復旧処理、画面更新の復帰を確認してください。
    On Error GoTo Finally

    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim wsSrc As Worksheet
    Set wsSrc = ActiveSheet

    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim keys As Object
    Set keys = CollectKeys(wsSrc, keyCol)

    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim lastRow As Long, lastCol As Long
    lastRow = GetLastRow(wsSrc)
    lastCol = GetLastCol(wsSrc)

    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim key As Variant
    Dim wsDest As Worksheet
    ' [ブロック解説: 繰り返し処理]
    ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
    For Each key In keys.keys
        ' [ブロック解説: オブジェクト参照の設定]
        ' Workbook、Worksheet、Range、外部COMオブジェクトなどの参照を初期化します。オブジェクトが存在し、必要な参照ライブラリまたはアプリケーションが利用可能か確認します。
        Set wsDest = GetOrCreateSheet(ThisWorkbook, SafeSheetName(CStr(key)))
        wsDest.Cells.Clear
        CopyFiltered wsSrc, keyCol, CStr(key), lastRow, lastCol, wsDest.Range("A1")
        wsDest.Columns.AutoFit
    Next key

    wsSrc.AutoFilterMode = False
    wsSrc.Activate

Finally:
    SpeedUp False
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If Err.Number <> 0 Then
        MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    Else
        MsgBox keys.Count & " 件のシートに振り分けました。", vbInformation
    End If
End Sub

'--------------------------------------------------
' メイン: キー列の値ごとに「別ブック」へ振り分け保存
'--------------------------------------------------
' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub 条件別ブック振り分け() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub 条件別ブック振り分け()
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim keyCol As Long
    keyCol = AskKeyColumn()
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If keyCol = 0 Then Exit Sub

    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim saveFolder As String
    saveFolder = SelectFolder("分割ブックの保存先フォルダを選択")
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If saveFolder = "" Then Exit Sub

    SpeedUp True
    ' [ブロック解説: エラー処理設定]
    ' エラー発生時の処理分岐を定義します。エラー番号、エラー内容、復旧処理、画面更新の復帰を確認してください。
    On Error GoTo Finally

    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim wsSrc As Worksheet
    Set wsSrc = ActiveSheet

    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim keys As Object
    Set keys = CollectKeys(wsSrc, keyCol)

    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim lastRow As Long, lastCol As Long
    lastRow = GetLastRow(wsSrc)
    lastCol = GetLastCol(wsSrc)

    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim key As Variant
    Dim wbNew As Workbook
    ' [ブロック解説: 繰り返し処理]
    ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
    For Each key In keys.keys
        ' [ブロック解説: オブジェクト参照の設定]
        ' Workbook、Worksheet、Range、外部COMオブジェクトなどの参照を初期化します。オブジェクトが存在し、必要な参照ライブラリまたはアプリケーションが利用可能か確認します。
        Set wbNew = Workbooks.Add(xlWBATWorksheet)   ' 1シートのみの新規ブック
        CopyFiltered wsSrc, keyCol, CStr(key), lastRow, lastCol, _
                     wbNew.Worksheets(1).Range("A1")
        wbNew.Worksheets(1).Name = SafeSheetName(CStr(key))
        wbNew.Worksheets(1).Columns.AutoFit
        wbNew.SaveAs _
            fileName:=saveFolder & Application.PathSeparator & SafeSheetName(CStr(key)) & ".xlsx", _
            FileFormat:=xlOpenXMLWorkbook
        wbNew.Close SaveChanges:=False
    Next key

    wsSrc.AutoFilterMode = False

Finally:
    SpeedUp False
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If Err.Number <> 0 Then
        MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    Else
        MsgBox keys.Count & " 件のブックを保存しました。" & vbCrLf & saveFolder, vbInformation
    End If
End Sub

'--------------------------------------------------
' キー列番号をユーザーに確認
'--------------------------------------------------
' [ブロック解説: プロシージャ]
' 宣言ブロックです。Private Function AskKeyColumn() As Long の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Private Function AskKeyColumn() As Long
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim ans As String
    ans = InputBox("振り分けキーの列番号を入力してください" & vbCrLf & _
                   "（例: A列=1, B列=2 ...）", "キー列の指定", "1")
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If ans = "" Or Not IsNumeric(ans) Then
        AskKeyColumn = 0
    Else
        AskKeyColumn = CLng(ans)
    End If
End Function

'--------------------------------------------------
' キー列のユニーク値を収集（Dictionary）
'--------------------------------------------------
' [ブロック解説: プロシージャ]
' 宣言ブロックです。Private Function CollectKeys(ByVal ws As Worksheet, ByVal keyCol As Long) As Object の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Private Function CollectKeys(ByVal ws As Worksheet, ByVal keyCol As Long) As Object
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim dic As Object
    ' [ブロック解説: オブジェクト参照の設定]
    ' Workbook、Worksheet、Range、外部COMオブジェクトなどの参照を初期化します。オブジェクトが存在し、必要な参照ライブラリまたはアプリケーションが利用可能か確認します。
    Set dic = CreateObject("Scripting.Dictionary")

    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim lastRow As Long, r As Long
    Dim v As String
    lastRow = GetLastRow(ws)
    ' [ブロック解説: 繰り返し処理]
    ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
    For r = 2 To lastRow
        v = Trim$(CStr(ws.Cells(r, keyCol).Value))
        ' [ブロック解説: 条件分岐]
        ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
        If v <> "" Then dic(v) = True
    Next r
    Set CollectKeys = dic
End Function

'--------------------------------------------------
' AutoFilterで抽出して貼り付け
'--------------------------------------------------
' [ブロック解説: プロシージャ]
' 宣言ブロックです。Private Sub CopyFiltered(ByVal wsSrc As Worksheet, ByVal keyCol As Long, _ の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Private Sub CopyFiltered(ByVal wsSrc As Worksheet, ByVal keyCol As Long, _
                         ByVal key As String, ByVal lastRow As Long, _
                         ByVal lastCol As Long, ByVal destCell As Range)
    ' [ブロック解説: 一括設定ブロック]
    ' Withブロック内のオブジェクトへ、セル形式、印刷設定、オプションなどをまとめて設定します。設定対象と上書き範囲に注意します。
    With wsSrc.Range(wsSrc.Cells(1, 1), wsSrc.Cells(lastRow, lastCol))
        .AutoFilter Field:=keyCol, Criteria1:=key
        .SpecialCells(xlCellTypeVisible).Copy destCell
    End With
    Application.CutCopyMode = False
    wsSrc.AutoFilterMode = False
End Sub

