Attribute VB_Name = "Mod112_リネームリスト"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod112_リネームリスト
' - ファイル名: Mod112_リネームリスト.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub 処理対象リストに基づき一括リネーム() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub 処理対象リストに基づき一括リネーム()
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If lastRow < 2 Then
        MsgBox "A列に変更前パス、B列に変更後パスを設定してください。", vbExclamation
        Exit Sub
    End If
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If MsgBox(lastRow - 1 & " 件をリネームしますか？", vbYesNo + vbQuestion) <> vbYes Then Exit Sub

    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim r As Long
    ' [ブロック解説: 繰り返し処理]
    ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
    For r = 2 To lastRow
        ' [ブロック解説: 変数・定数・設定値の宣言]
        ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
        Dim sourcePath As String: sourcePath = CStr(ws.Cells(r, 1).Value)
        Dim destinationPath As String: destinationPath = CStr(ws.Cells(r, 2).Value)
        ' [ブロック解説: 条件分岐]
        ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
        If fso.FileExists(sourcePath) And destinationPath <> "" Then
            ' [ブロック解説: 条件分岐]
            ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
            If fso.FileExists(destinationPath) Then
                ws.Cells(r, 3).Value = "移動先が既に存在"
            Else
                fso.MoveFile sourcePath, destinationPath
                ws.Cells(r, 3).Value = "完了"
            End If
        Else
            ws.Cells(r, 3).Value = "元ファイルなし"
        End If
    Next r
End Sub

