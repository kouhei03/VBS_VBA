Attribute VB_Name = "Mod29_シート操作"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod29_シート操作
' - ファイル名: Mod29_シート操作.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Function シート追加(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Function シート追加(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    ' [ブロック解説: オブジェクト参照の設定]
    ' Workbook、Worksheet、Range、外部COMオブジェクトなどの参照を初期化します。オブジェクトが存在し、必要な参照ライブラリまたはアプリケーションが利用可能か確認します。
    Set シート追加 = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    シート追加.Name = sheetName
End Function

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub シート名称変更(ByVal ws As Worksheet, ByVal newName As String) の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub シート名称変更(ByVal ws As Worksheet, ByVal newName As String)
    ws.Name = newName
End Sub

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub シート削除(ByVal ws As Worksheet) の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub シート削除(ByVal ws As Worksheet)
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If wbSheetCount(ws.Parent) <= 1 Then
        Err.Raise vbObjectError + 2901, , "ブック内の最後のシートは削除できません。"
    End If
    ' [ブロック解説: Excel実行環境の設定]
    ' 画面更新、計算モード、アラート、イベントの実行設定を変更します。処理終了時に元の設定へ復帰できることを確認してください。
    Application.DisplayAlerts = False
    ws.Delete
    ' [ブロック解説: Excel実行環境の設定]
    ' 画面更新、計算モード、アラート、イベントの実行設定を変更します。処理終了時に元の設定へ復帰できることを確認してください。
    Application.DisplayAlerts = True
End Sub

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Private Function wbSheetCount(ByVal wb As Workbook) As Long の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Private Function wbSheetCount(ByVal wb As Workbook) As Long
    wbSheetCount = wb.Worksheets.Count
End Function

