Attribute VB_Name = "Mod40_暗号化圧縮補助"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod40_暗号化圧縮補助
' - ファイル名: Mod40_暗号化圧縮補助.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub 外部連携データ暗号化圧縮Zip前処理補助() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub 外部連携データ暗号化圧縮Zip前処理補助()
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim sourceFolder As String
    ' [ブロック解説: 一括設定ブロック]
    ' Withブロック内のオブジェクトへ、セル形式、印刷設定、オプションなどをまとめて設定します。設定対象と上書き範囲に注意します。
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "圧縮する連携ファイルがあるフォルダを選択"
        ' [ブロック解説: 条件分岐]
        ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
        If .Show <> -1 Then Exit Sub
        sourceFolder = .SelectedItems(1)
    End With
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim zipPath As String
    zipPath = Application.GetSaveAsFilename("連携データ.zip", "Zipファイル (*.zip),*.zip")
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If zipPath = "False" Then Exit Sub
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim shell As Object: Set shell = CreateObject("Shell.Application")
    Dim zipFolder As Object
    CreateEmptyZip zipPath
    Set zipFolder = shell.NameSpace(zipPath)
    zipFolder.CopyHere shell.NameSpace(sourceFolder).Items, 16
    MsgBox "圧縮処理を開始しました。" & vbCrLf & zipPath, vbInformation
End Sub

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Private Sub CreateEmptyZip(ByVal zipPath As String) の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Private Sub CreateEmptyZip(ByVal zipPath As String)
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim fileNumber As Integer: fileNumber = FreeFile
    Open zipPath For Output As #fileNumber
    Print #fileNumber, "PK" & Chr$(5) & Chr$(6) & String$(18, Chr$(0))
    Close #fileNumber
End Sub

