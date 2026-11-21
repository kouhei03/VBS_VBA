Attribute VB_Name = "Mod98_CSV読み込み"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod98_CSV読み込み
' - ファイル名: Mod98_CSV読み込み.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub CSV読み込み文字コード対応() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub CSV読み込み文字コード対応()
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim filePath As String
    ' [ブロック解説: 一括設定ブロック]
    ' Withブロック内のオブジェクトへ、セル形式、印刷設定、オプションなどをまとめて設定します。設定対象と上書き範囲に注意します。
    With Application.FileDialog(msoFileDialogFilePicker)
        .Filters.Clear
        .Filters.Add "CSV", "*.csv"
        ' [ブロック解説: 条件分岐]
        ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
        If .Show <> -1 Then Exit Sub
        filePath = .SelectedItems(1)
    End With
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim codePage As String
    codePage = InputBox("文字コードを入力: 65001=UTF-8、932=Shift_JIS、1252=ANSI", "CSV読込", "65001")
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If codePage = "65001" Then
        ' [ブロック解説: 変数・定数・設定値の宣言]
        ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
        Dim text As String: text = ReadUtf8(filePath)
        Dim data As Variant: data = Split(text, vbCrLf)
        Dim i As Long
        Dim ws As Worksheet: Set ws = ActiveSheet
        ' [ブロック解説: 繰り返し処理]
        ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
        For i = LBound(data) To UBound(data)
            ws.Cells(i + 1, 1).Value = data(i)
        Next i
    Else
        Workbooks.OpenText fileName:=filePath, Origin:=CLng(codePage), StartRow:=1, _
                           DataType:=xlDelimited, TextQualifier:=xlTextQualifierDoubleQuote, _
                           Comma:=True, Local:=True
    End If
End Sub

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Private Function ReadUtf8(ByVal filePath As String) As String の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Private Function ReadUtf8(ByVal filePath As String) As String
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim stream As Object: Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile filePath
    ReadUtf8 = stream.ReadText
    stream.Close
End Function

