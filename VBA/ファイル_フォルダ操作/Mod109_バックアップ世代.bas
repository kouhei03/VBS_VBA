Attribute VB_Name = "Mod109_バックアップ世代"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod109_バックアップ世代
' - ファイル名: Mod109_バックアップ世代.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub バックアップフォルダへ世代管理保存() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub バックアップフォルダへ世代管理保存()
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If Not fso.FileExists(ThisWorkbook.FullName) Then
        MsgBox "保存済みブックに対して実行してください。", vbExclamation
        Exit Sub
    End If
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim backupParent As String
    ' [ブロック解説: 一括設定ブロック]
    ' Withブロック内のオブジェクトへ、セル形式、印刷設定、オプションなどをまとめて設定します。設定対象と上書き範囲に注意します。
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "バックアップ先フォルダを選択"
        ' [ブロック解説: 条件分岐]
        ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
        If .Show <> -1 Then Exit Sub
        backupParent = .SelectedItems(1)
    End With
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim generationFolder As String
    generationFolder = fso.BuildPath(backupParent, Format$(Now, "yyyymmdd_hhnnss"))
    fso.CreateFolder generationFolder
    fso.CopyFile ThisWorkbook.FullName, fso.BuildPath(generationFolder, fso.GetFileName(ThisWorkbook.FullName)), True
    MsgBox "世代バックアップを保存しました。" & vbCrLf & generationFolder, vbInformation
End Sub

