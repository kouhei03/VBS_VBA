Attribute VB_Name = "Mod117_ドライブ容量"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod117_ドライブ容量
' - ファイル名: Mod117_ドライブ容量.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub ドライブ空き容量チェック() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub ドライブ空き容量チェック()
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim driveLetter As String
    driveLetter = InputBox("確認するドライブ名を入力（例: C:）", "ドライブ容量", "C:")
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If Trim$(driveLetter) = "" Then Exit Sub
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If Not fso.DriveExists(driveLetter) Then
        MsgBox "ドライブが見つかりません。", vbCritical
        Exit Sub
    End If
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim drive As Object: Set drive = fso.GetDrive(driveLetter)
    Dim thresholdGB As Variant
    thresholdGB = Application.InputBox("警告する空き容量（GB）", "ドライブ容量", 10, , , , , 1)
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If thresholdGB = False Or Not IsNumeric(thresholdGB) Then Exit Sub
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim freeGB As Double: freeGB = CDbl(drive.FreeSpace) / 1024 / 1024 / 1024
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If freeGB < CDbl(thresholdGB) Then
        MsgBox "空き容量が少ない可能性があります。" & vbCrLf & Format$(freeGB, "0.00") & " GB", vbCritical
    Else
        MsgBox "空き容量: " & Format$(freeGB, "0.00") & " GB", vbInformation
    End If
End Sub

