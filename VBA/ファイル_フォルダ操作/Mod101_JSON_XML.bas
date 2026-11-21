Attribute VB_Name = "Mod101_JSON_XML"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod101_JSON_XML
' - ファイル名: Mod101_JSON_XML.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub JSON_XML読み込みパース() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub JSON_XML読み込みパース()
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim filePath As String
    ' [ブロック解説: 一括設定ブロック]
    ' Withブロック内のオブジェクトへ、セル形式、印刷設定、オプションなどをまとめて設定します。設定対象と上書き範囲に注意します。
    With Application.FileDialog(msoFileDialogFilePicker)
        .Filters.Clear
        .Filters.Add "JSON/XML", "*.json;*.xml"
        ' [ブロック解説: 条件分岐]
        ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
        If .Show <> -1 Then Exit Sub
        filePath = .SelectedItems(1)
    End With

    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim extension As String
    extension = LCase$(Right$(filePath, 5))
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim ws As Worksheet: Set ws = ActiveSheet
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If extension Like "*.json" Then
        ws.Cells(1, 1).Value = "JSONテキスト"
        ws.Cells(2, 1).Value = ReadUtf8(filePath)
    Else
        ' [ブロック解説: 変数・定数・設定値の宣言]
        ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
        Dim xml As Object: Set xml = CreateObject("MSXML2.DOMDocument")
        xml.Async = False
        ' [ブロック解説: 条件分岐]
        ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
        If Not xml.Load(filePath) Then
            MsgBox "XMLの解析に失敗しました。" & vbCrLf & xml.ParseError.Reason, vbCritical
            Exit Sub
        End If
        ws.Cells(1, 1).Value = "ノード名"
        ws.Cells(1, 2).Value = "値"
        WriteXmlNodes xml.DocumentElement, ws, 2
    End If
End Sub

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Private Sub WriteXmlNodes(ByVal node As Object, ByVal ws As Worksheet, ByRef r As Long) の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Private Sub WriteXmlNodes(ByVal node As Object, ByVal ws As Worksheet, ByRef r As Long)
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim child As Object
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If node.ChildNodes.Length > 0 And node.FirstChild.NodeType = 1 Then
        ' [ブロック解説: 繰り返し処理]
        ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
        For Each child In node.ChildNodes
            WriteXmlNodes child, ws, r
        Next child
    Else
        ws.Cells(r, 1).Value = node.BaseName
        ws.Cells(r, 2).Value = node.Text
        r = r + 1
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

