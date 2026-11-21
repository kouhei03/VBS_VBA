Attribute VB_Name = "Mod195_在庫発注点"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod195_在庫発注点
' - ファイル名: Mod195_在庫発注点.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub 在庫一覧発注点アラート帳票作成() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub 在庫一覧発注点アラート帳票作成()
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim source As Worksheet: Set source = ActiveSheet
    Dim output As Worksheet: Set output = GetOrCreateSheet("在庫一覧")
    output.Cells.Clear
    source.UsedRange.Copy output.Range("A1")
    output.Cells(1, 4).Value = "発注点判定"
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim lastRow As Long: lastRow = output.Cells(output.Rows.Count, 1).End(xlUp).Row
    Dim r As Long
    ' [ブロック解説: 繰り返し処理]
    ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
    For r = 2 To lastRow
        ' [ブロック解説: 条件分岐]
        ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
        If IsNumeric(output.Cells(r, 2).Value) And IsNumeric(output.Cells(r, 3).Value) Then
            ' [ブロック解説: 条件分岐]
            ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
            If output.Cells(r, 2).Value <= output.Cells(r, 3).Value Then
                output.Cells(r, 4).Value = "発注"
                output.Rows(r).Interior.Color = RGB(255, 199, 206)
            Else
                output.Cells(r, 4).Value = "十分"
            End If
        End If
    Next r
    output.Columns.AutoFit
End Sub

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Private Function GetOrCreateSheet(ByVal name As String) As Worksheet の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Private Function GetOrCreateSheet(ByVal name As String) As Worksheet
    ' [ブロック解説: エラー処理設定]
    ' エラー発生時の処理分岐を定義します。エラー番号、エラー内容、復旧処理、画面更新の復帰を確認してください。
    On Error Resume Next
    ' [ブロック解説: オブジェクト参照の設定]
    ' Workbook、Worksheet、Range、外部COMオブジェクトなどの参照を初期化します。オブジェクトが存在し、必要な参照ライブラリまたはアプリケーションが利用可能か確認します。
    Set GetOrCreateSheet = ThisWorkbook.Worksheets(name)
    ' [ブロック解説: エラー処理設定]
    ' エラー発生時の処理分岐を定義します。エラー番号、エラー内容、復旧処理、画面更新の復帰を確認してください。
    On Error GoTo 0
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If GetOrCreateSheet Is Nothing Then
        ' [ブロック解説: オブジェクト参照の設定]
        ' Workbook、Worksheet、Range、外部COMオブジェクトなどの参照を初期化します。オブジェクトが存在し、必要な参照ライブラリまたはアプリケーションが利用可能か確認します。
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = name
    End If
End Function

