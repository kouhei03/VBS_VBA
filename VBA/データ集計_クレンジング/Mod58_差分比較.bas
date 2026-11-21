Attribute VB_Name = "Mod58_差分比較"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod58_差分比較
' - ファイル名: Mod58_差分比較.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub 新旧データの差分をハイライト() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub 新旧データの差分をハイライト()
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim oldName As String: oldName = InputBox("旧データシート名", "差分比較")
    Dim newName As String: newName = InputBox("新データシート名", "差分比較")
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If Not SheetExists(ThisWorkbook, oldName) Or Not SheetExists(ThisWorkbook, newName) Then Exit Sub
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim oldWs As Worksheet: Set oldWs = ThisWorkbook.Worksheets(oldName)
    Dim newWs As Worksheet: Set newWs = ThisWorkbook.Worksheets(newName)
    Dim lastRow As Long: lastRow = Application.WorksheetFunction.Max(GetLastRow(oldWs), GetLastRow(newWs))
    Dim lastCol As Long: lastCol = Application.WorksheetFunction.Max(GetLastCol(oldWs), GetLastCol(newWs))
    Dim r As Long, c As Long
    ' [ブロック解説: 繰り返し処理]
    ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
    For r = 1 To lastRow
        ' [ブロック解説: 繰り返し処理]
        ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
        For c = 1 To lastCol
            ' [ブロック解説: 条件分岐]
            ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
            If CStr(oldWs.Cells(r, c).Value) <> CStr(newWs.Cells(r, c).Value) Then
                newWs.Cells(r, c).Interior.Color = RGB(255, 230, 153)
            Else
                newWs.Cells(r, c).Interior.Pattern = xlNone
            End If
        Next c
    Next r
End Sub

