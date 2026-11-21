Attribute VB_Name = "Mod04_重複処理"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod04_重複処理
' - ファイル名: Mod04_重複処理.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

'==================================================
' Mod04_重複処理
' 顧客リスト・コード一覧などから重複行を検出し、
' ハイライト（色付け） または 削除 する
'
' 前提: アクティブシートは1行目がヘッダーの表形式
'==================================================

' [ブロック解説: 変数・定数・設定値の宣言]
' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
Private Const HIGHLIGHT_COLOR As Long = &H9AFFFF   ' 薄い黄色 (RGB 255,255,154相当)

'--------------------------------------------------
' メイン: 重複行をハイライト（2件目以降に色付け）
'--------------------------------------------------
' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub 重複ハイライト() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub 重複ハイライト()
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim keyCols As Variant
    keyCols = AskKeyColumns()
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If IsEmpty(keyCols) Then Exit Sub

    SpeedUp True
    ' [ブロック解説: エラー処理設定]
    ' エラー発生時の処理分岐を定義します。エラー番号、エラー内容、復旧処理、画面更新の復帰を確認してください。
    On Error GoTo Finally

    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim lastRow As Long, lastCol As Long
    lastRow = GetLastRow(ws)
    lastCol = GetLastCol(ws)

    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim dic As Object
    ' [ブロック解説: オブジェクト参照の設定]
    ' Workbook、Worksheet、Range、外部COMオブジェクトなどの参照を初期化します。オブジェクトが存在し、必要な参照ライブラリまたはアプリケーションが利用可能か確認します。
    Set dic = CreateObject("Scripting.Dictionary")

    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim r As Long
    Dim key As String
    Dim dupCount As Long
    ' [ブロック解説: 繰り返し処理]
    ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
    For r = 2 To lastRow
        key = BuildKey(ws, r, keyCols)
        ' [ブロック解説: 条件分岐]
        ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
        If dic.Exists(key) Then
            ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol)).Interior.Color = HIGHLIGHT_COLOR
            dupCount = dupCount + 1
        Else
            dic(key) = r
        End If
    Next r

Finally:
    SpeedUp False
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If Err.Number <> 0 Then
        MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    Else
        MsgBox "重複行: " & dupCount & " 行をハイライトしました。", vbInformation
    End If
End Sub

'--------------------------------------------------
' メイン: 重複行を削除（最初の1件を残す）
'--------------------------------------------------
' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub 重複削除() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub 重複削除()
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim keyCols As Variant
    keyCols = AskKeyColumns()
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If IsEmpty(keyCols) Then Exit Sub

    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If MsgBox("重複行を削除します（最初の1件は残します）。" & vbCrLf & _
              "元に戻せないため、事前にバックアップを推奨します。" & vbCrLf & vbCrLf & _
              "実行しますか？", vbYesNo + vbExclamation) <> vbYes Then Exit Sub

    SpeedUp True
    ' [ブロック解説: エラー処理設定]
    ' エラー発生時の処理分岐を定義します。エラー番号、エラー内容、復旧処理、画面更新の復帰を確認してください。
    On Error GoTo Finally

    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim lastRow As Long
    lastRow = GetLastRow(ws)

    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim dic As Object
    ' [ブロック解説: オブジェクト参照の設定]
    ' Workbook、Worksheet、Range、外部COMオブジェクトなどの参照を初期化します。オブジェクトが存在し、必要な参照ライブラリまたはアプリケーションが利用可能か確認します。
    Set dic = CreateObject("Scripting.Dictionary")

    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim delRange As Range
    Dim r As Long
    Dim key As String
    Dim dupCount As Long
    ' [ブロック解説: 繰り返し処理]
    ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
    For r = 2 To lastRow
        key = BuildKey(ws, r, keyCols)
        ' [ブロック解説: 条件分岐]
        ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
        If dic.Exists(key) Then
            ' [ブロック解説: 条件分岐]
            ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
            If delRange Is Nothing Then
                ' [ブロック解説: オブジェクト参照の設定]
                ' Workbook、Worksheet、Range、外部COMオブジェクトなどの参照を初期化します。オブジェクトが存在し、必要な参照ライブラリまたはアプリケーションが利用可能か確認します。
                Set delRange = ws.Rows(r)
            Else
                ' [ブロック解説: オブジェクト参照の設定]
                ' Workbook、Worksheet、Range、外部COMオブジェクトなどの参照を初期化します。オブジェクトが存在し、必要な参照ライブラリまたはアプリケーションが利用可能か確認します。
                Set delRange = Union(delRange, ws.Rows(r))
            End If
            dupCount = dupCount + 1
        Else
            dic(key) = r
        End If
    Next r

    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If Not delRange Is Nothing Then delRange.Delete

Finally:
    SpeedUp False
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If Err.Number <> 0 Then
        MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    Else
        MsgBox "重複行: " & dupCount & " 行を削除しました。", vbInformation
    End If
End Sub

'--------------------------------------------------
' ハイライトを解除
'--------------------------------------------------
' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub ハイライト解除() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub ハイライト解除()
    ActiveSheet.UsedRange.Interior.ColorIndex = xlNone
    MsgBox "ハイライトを解除しました。", vbInformation
End Sub

'--------------------------------------------------
' 重複判定キーの列番号をユーザーに確認
' 複数列はカンマ区切り（例: 1,3）で複合キー判定
'--------------------------------------------------
' [ブロック解説: プロシージャ]
' 宣言ブロックです。Private Function AskKeyColumns() As Variant の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Private Function AskKeyColumns() As Variant
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim ans As String
    ans = InputBox("重複判定に使う列番号を入力してください" & vbCrLf & _
                   "（例: A列のみ → 1 / A列+C列の複合キー → 1,3）", _
                   "キー列の指定", "1")
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If ans = "" Then
        AskKeyColumns = Empty
        Exit Function
    End If

    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim parts As Variant
    parts = Split(ans, ",")

    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim i As Long
    Dim cols() As Long
    ReDim cols(LBound(parts) To UBound(parts))
    ' [ブロック解説: 繰り返し処理]
    ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
    For i = LBound(parts) To UBound(parts)
        ' [ブロック解説: 条件分岐]
        ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
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
' [ブロック解説: プロシージャ]
' 宣言ブロックです。Private Function BuildKey(ByVal ws As Worksheet, ByVal r As Long, _ の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Private Function BuildKey(ByVal ws As Worksheet, ByVal r As Long, _
                          ByVal keyCols As Variant) As String
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim i As Long
    Dim key As String
    ' [ブロック解説: 繰り返し処理]
    ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
    For i = LBound(keyCols) To UBound(keyCols)
        key = key & CleanText(CStr(ws.Cells(r, keyCols(i)).Value)) & vbNullChar
    Next i
    BuildKey = key
End Function

