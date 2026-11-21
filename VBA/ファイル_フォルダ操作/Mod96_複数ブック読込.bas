Attribute VB_Name = "Mod96_複数ブック読込"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod96_複数ブック読込
' - ファイル名: Mod96_複数ブック読込.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub 複数Excelブック順次オープンと取得() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub 複数Excelブック順次オープンと取得()
    ' [ブロック解説: 一括設定ブロック]
    ' Withブロック内のオブジェクトへ、セル形式、印刷設定、オプションなどをまとめて設定します。設定対象と上書き範囲に注意します。
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "Excel", "*.xls;*.xlsx;*.xlsm;*.xlsb"
        ' [ブロック解説: 条件分岐]
        ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
        If .Show <> -1 Then Exit Sub

        ' [ブロック解説: 変数・定数・設定値の宣言]
        ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
        Dim output As Worksheet: Set output = GetOrCreateSheet("複数ブック読込")
        output.Cells.Clear
        output.Cells(1, 1).Value = "ブック"
        output.Cells(1, 2).Value = "シート"
        output.Cells(1, 3).Value = "最終行"
        output.Cells(1, 4).Value = "最終列"

        ' [ブロック解説: 変数・定数・設定値の宣言]
        ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
        Dim i As Long, r As Long: r = 2
        Dim source As Workbook
        ' [ブロック解説: 繰り返し処理]
        ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
        For i = 1 To .SelectedItems.Count
            ' [ブロック解説: オブジェクト参照の設定]
            ' Workbook、Worksheet、Range、外部COMオブジェクトなどの参照を初期化します。オブジェクトが存在し、必要な参照ライブラリまたはアプリケーションが利用可能か確認します。
            Set source = Workbooks.Open(.SelectedItems(i), UpdateLinks:=0, ReadOnly:=True)
            ' [ブロック解説: 変数・定数・設定値の宣言]
            ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
            Dim ws As Worksheet
            ' [ブロック解説: 繰り返し処理]
            ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
            For Each ws In source.Worksheets
                output.Cells(r, 1).Value = source.Name
                output.Cells(r, 2).Value = ws.Name
                output.Cells(r, 3).Value = ws.Cells.Find("*", , xlFormulas, , xlByRows, xlPrevious).Row
                output.Cells(r, 4).Value = ws.Cells.Find("*", , xlFormulas, , xlByColumns, xlPrevious).Column
                r = r + 1
            Next ws
            source.Close False
        Next i
        output.Columns("A:D").AutoFit
    End With
End Sub

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Private Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Private Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet
    ' [ブロック解説: エラー処理設定]
    ' エラー発生時の処理分岐を定義します。エラー番号、エラー内容、復旧処理、画面更新の復帰を確認してください。
    On Error Resume Next
    ' [ブロック解説: オブジェクト参照の設定]
    ' Workbook、Worksheet、Range、外部COMオブジェクトなどの参照を初期化します。オブジェクトが存在し、必要な参照ライブラリまたはアプリケーションが利用可能か確認します。
    Set GetOrCreateSheet = ThisWorkbook.Worksheets(sheetName)
    ' [ブロック解説: エラー処理設定]
    ' エラー発生時の処理分岐を定義します。エラー番号、エラー内容、復旧処理、画面更新の復帰を確認してください。
    On Error GoTo 0
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If GetOrCreateSheet Is Nothing Then
        ' [ブロック解説: オブジェクト参照の設定]
        ' Workbook、Worksheet、Range、外部COMオブジェクトなどの参照を初期化します。オブジェクトが存在し、必要な参照ライブラリまたはアプリケーションが利用可能か確認します。
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = sheetName
    End If
End Function

