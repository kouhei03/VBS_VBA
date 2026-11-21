Attribute VB_Name = "Mod57_マスタ照合補完"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod57_マスタ照合補完
' - ファイル名: Mod57_マスタ照合補完.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub マスタデータを照合して補完() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub マスタデータを照合して補完()
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim targetName As String: targetName = InputBox("対象シート名", "マスタ照合")
    Dim masterName As String: masterName = InputBox("マスタシート名", "マスタ照合")
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If Not SheetExists(ThisWorkbook, targetName) Or Not SheetExists(ThisWorkbook, masterName) Then Exit Sub
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim targetWs As Worksheet: Set targetWs = ThisWorkbook.Worksheets(targetName)
    Dim masterWs As Worksheet: Set masterWs = ThisWorkbook.Worksheets(masterName)
    Dim masterMap As Object: Set masterMap = CreateObject("Scripting.Dictionary")
    Dim r As Long
    ' [ブロック解説: 繰り返し処理]
    ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
    For r = 2 To GetLastRow(masterWs)
        masterMap(CStr(masterWs.Cells(r, 1).Value)) = masterWs.Cells(r, 2).Value
    Next r
    ' [ブロック解説: 繰り返し処理]
    ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
    For r = 2 To GetLastRow(targetWs)
        ' [ブロック解説: 変数・定数・設定値の宣言]
        ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
        Dim key As String: key = CStr(targetWs.Cells(r, 1).Value)
        ' [ブロック解説: 条件分岐]
        ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
        If Len(CStr(targetWs.Cells(r, 2).Value)) = 0 And masterMap.Exists(key) Then
            targetWs.Cells(r, 2).Value = masterMap(key)
        End If
    Next r
End Sub

