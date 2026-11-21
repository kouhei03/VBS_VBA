Attribute VB_Name = "Mod104_シート個別保存"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod104_シート個別保存
' - ファイル名: Mod104_シート個別保存.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub 複数シートを個別ブックへ分割保存() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub 複数シートを個別ブックへ分割保存()
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim destinationFolder As String
    ' [ブロック解説: 一括設定ブロック]
    ' Withブロック内のオブジェクトへ、セル形式、印刷設定、オプションなどをまとめて設定します。設定対象と上書き範囲に注意します。
    With Application.FileDialog(msoFileDialogFolderPicker)
        ' [ブロック解説: 条件分岐]
        ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
        If .Show <> -1 Then Exit Sub
        destinationFolder = .SelectedItems(1)
    End With
    ' [ブロック解説: Excel実行環境の設定]
    ' 画面更新、計算モード、アラート、イベントの実行設定を変更します。処理終了時に元の設定へ復帰できることを確認してください。
    Application.DisplayAlerts = False
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim ws As Worksheet
    ' [ブロック解説: 繰り返し処理]
    ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
    For Each ws In ThisWorkbook.Worksheets
        ' [ブロック解説: 変数・定数・設定値の宣言]
        ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
        Dim temp As Workbook
        ' [ブロック解説: オブジェクト参照の設定]
        ' Workbook、Worksheet、Range、外部COMオブジェクトなどの参照を初期化します。オブジェクトが存在し、必要な参照ライブラリまたはアプリケーションが利用可能か確認します。
        Set temp = Workbooks.Add(1)
        ws.Copy Before:=temp.Worksheets(1)
        temp.Worksheets(1).Delete
        temp.SaveAs destinationFolder & "\" & SafeSheetName(ws.Name) & ".xlsx", xlOpenXMLWorkbook
        temp.Close False
    Next ws
    ' [ブロック解説: Excel実行環境の設定]
    ' 画面更新、計算モード、アラート、イベントの実行設定を変更します。処理終了時に元の設定へ復帰できることを確認してください。
    Application.DisplayAlerts = True
    MsgBox "シートを個別ブックに分割保存しました。", vbInformation
End Sub

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Private Function SafeSheetName(ByVal name As String) As String の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Private Function SafeSheetName(ByVal name As String) As String
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim invalid As Variant: invalid = Array("\", "/", ":", "*", "?", Chr$(34), "<", ">", "|")
    Dim i As Long
    ' [ブロック解説: 繰り返し処理]
    ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
    For i = LBound(invalid) To UBound(invalid)
        name = Replace(name, invalid(i), "_")
    Next i
    SafeSheetName = name
End Function

