Attribute VB_Name = "Mod30_シート表示制御"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod30_シート表示制御
' - ファイル名: Mod30_シート表示制御.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub シート表示(ByVal ws As Worksheet) の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub シート表示(ByVal ws As Worksheet)
    ws.Visible = xlSheetVisible
End Sub

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub シート非表示(ByVal ws As Worksheet) の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub シート非表示(ByVal ws As Worksheet)
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If wbVisibleSheetCount(ws.Parent) <= 1 Then
        Err.Raise vbObjectError + 3001, , "表示中の最後のシートは非表示にできません。"
    End If
    ws.Visible = xlSheetHidden
End Sub

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub シートを非常に隠す(ByVal ws As Worksheet) の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub シートを非常に隠す(ByVal ws As Worksheet)
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If wbVisibleSheetCount(ws.Parent) <= 1 Then
        Err.Raise vbObjectError + 3002, , "表示中の最後のシートをVeryHiddenにはできません。"
    End If
    ws.Visible = xlSheetVeryHidden
End Sub

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Private Function wbVisibleSheetCount(ByVal wb As Workbook) As Long の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Private Function wbVisibleSheetCount(ByVal wb As Workbook) As Long
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim ws As Worksheet
    ' [ブロック解説: 繰り返し処理]
    ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
    For Each ws In wb.Worksheets
        ' [ブロック解説: 条件分岐]
        ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
        If ws.Visible = xlSheetVisible Then wbVisibleSheetCount = wbVisibleSheetCount + 1
    Next ws
End Function

