Attribute VB_Name = "Mod110_古いバックアップ削除"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod110_古いバックアップ削除
' - ファイル名: Mod110_古いバックアップ削除.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub 古いバックアップファイル自動削除() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub 古いバックアップファイル自動削除()
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim backupFolder As String
    ' [ブロック解説: 一括設定ブロック]
    ' Withブロック内のオブジェクトへ、セル形式、印刷設定、オプションなどをまとめて設定します。設定対象と上書き範囲に注意します。
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "バックアップフォルダを選択"
        ' [ブロック解説: 条件分岐]
        ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
        If .Show <> -1 Then Exit Sub
        backupFolder = .SelectedItems(1)
    End With
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim keepDays As Variant
    keepDays = Application.InputBox("保持日数を入力してください", "世代管理", 30, , , , , 1)
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If keepDays = False Or Not IsNumeric(keepDays) Then Exit Sub

    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim deletedCount As Long: deletedCount = 0
    Dim cutoff As Date: cutoff = DateAdd("d", -CLng(keepDays), Now)
    Dim file As Object
    ' [ブロック解説: 繰り返し処理]
    ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
    For Each file In fso.GetFolder(backupFolder).Files
        ' [ブロック解説: 条件分岐]
        ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
        If file.DateLastModified < cutoff Then
            file.Delete True
            deletedCount = deletedCount + 1
        End If
    Next file
    MsgBox deletedCount & " 件削除しました。", vbInformation
End Sub

