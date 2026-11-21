Attribute VB_Name = "Mod109_チケット集計"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod109_チケット集計
' - ファイル名: Mod109_チケット集計.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub チケット対応状況平均解決時間集計() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub チケット対応状況平均解決時間集計()
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim count As Long, totalHours As Double
    Dim r As Long
    ' [ブロック解説: 繰り返し処理]
    ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        ' [ブロック解説: 条件分岐]
        ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
        If CStr(ws.Cells(r, 4).Value) = "完了" _
           And IsDate(ws.Cells(r, 2).Value) And IsDate(ws.Cells(r, 3).Value) Then
            count = count + 1
            totalHours = totalHours + CDbl(CDate(ws.Cells(r, 3).Value) - CDate(ws.Cells(r, 2).Value)) * 24
        End If
    Next r
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim avg As Double: avg = 0
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If count > 0 Then avg = totalHours / count
    MsgBox "完了件数: " & count & vbCrLf & "平均解決時間: " & Format$(avg, "0.0") & " 時間", vbInformation
End Sub

