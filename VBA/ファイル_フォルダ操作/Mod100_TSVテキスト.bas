Attribute VB_Name = "Mod100_TSVテキスト"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod100_TSVテキスト
' - ファイル名: Mod100_TSVテキスト.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub TSVテキスト読み書き() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub TSVテキスト読み書き()
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim mode As String
    mode = InputBox("処理を入力: 1=読み込み、2=書き出し", "TSV/テキスト", "1")
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If mode = "1" Then
        ' [ブロック解説: 変数・定数・設定値の宣言]
        ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
        Dim filePath As String
        ' [ブロック解説: 一括設定ブロック]
        ' Withブロック内のオブジェクトへ、セル形式、印刷設定、オプションなどをまとめて設定します。設定対象と上書き範囲に注意します。
        With Application.FileDialog(msoFileDialogFilePicker)
            .Filters.Add "Text", "*.txt;*.tsv"
            ' [ブロック解説: 条件分岐]
            ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
            If .Show <> -1 Then Exit Sub
            filePath = .SelectedItems(1)
        End With
        ' [ブロック解説: 変数・定数・設定値の宣言]
        ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
        Dim separator As String: separator = IIf(LCase$(Right$(filePath, 4)) = ".tsv", vbTab, vbLf)
        Dim text As String
        text = ReadText(filePath)
        ActiveSheet.Cells(1, 1).Value = text
    ElseIf mode = "2" Then
        ' [ブロック解説: 変数・定数・設定値の宣言]
        ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
        Dim destination As String
        ' [ブロック解説: 一括設定ブロック]
        ' Withブロック内のオブジェクトへ、セル形式、印刷設定、オプションなどをまとめて設定します。設定対象と上書き範囲に注意します。
        With Application.FileDialog(msoFileDialogSaveAs)
            .Filters.Clear
            .Filters.Add "Text", "*.txt"
            .Filters.Add "TSV", "*.tsv"
            ' [ブロック解説: 条件分岐]
            ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
            If .Show = -1 Then destination = .SelectedItems(1)
        End With
        ' [ブロック解説: 条件分岐]
        ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
        If destination <> "" Then WriteText destination, Selection.Text
    End If
End Sub

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Private Function ReadText(ByVal filePath As String) As String の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Private Function ReadText(ByVal filePath As String) As String
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim stream As Object: Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile filePath
    ReadText = stream.ReadText
    stream.Close
End Function

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Private Sub WriteText(ByVal filePath As String, ByVal text As String) の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Private Sub WriteText(ByVal filePath As String, ByVal text As String)
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim stream As Object: Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText text
    stream.SaveToFile filePath, 2
    stream.Close
End Sub

