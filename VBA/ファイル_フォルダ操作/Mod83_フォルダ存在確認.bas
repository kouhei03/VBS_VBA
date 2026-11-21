Attribute VB_Name = "Mod83_フォルダ存在確認"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod83_フォルダ存在確認
' - ファイル名: Mod83_フォルダ存在確認.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub フォルダ存在確認と作成() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub フォルダ存在確認と作成()
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim folderPath As String
    folderPath = InputBox("確認・作成するフォルダのフルパスを入力してください", "フォルダ作成")
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If Trim$(folderPath) = "" Then Exit Sub

    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If fso.FolderExists(folderPath) Then
        MsgBox "フォルダは既に存在します。" & vbCrLf & folderPath, vbInformation
    Else
        ' [ブロック解説: 条件分岐]
        ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
        If MsgBox("フォルダを作成しますか？" & vbCrLf & folderPath, vbYesNo + vbQuestion) = vbYes Then
            CreateFolderTree fso, folderPath
            MsgBox "フォルダを作成しました。", vbInformation
        End If
    End If
End Sub

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Private Sub CreateFolderTree(ByVal fso As Object, ByVal folderPath As String) の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Private Sub CreateFolderTree(ByVal fso As Object, ByVal folderPath As String)
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If fso.FolderExists(folderPath) Then Exit Sub
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim parentPath As String
    parentPath = fso.GetParentFolderName(folderPath)
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If parentPath <> "" And Not fso.FolderExists(parentPath) Then CreateFolderTree fso, parentPath
    fso.CreateFolder folderPath
End Sub

