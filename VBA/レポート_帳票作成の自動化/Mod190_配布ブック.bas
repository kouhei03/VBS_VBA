Attribute VB_Name = "Mod190_配布ブック"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod190_配布ブック
' - ファイル名: Mod190_配布ブック.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub 配布用ブック数式値化不要シート削除入力欄ロック() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub 配布用ブック数式値化不要シート削除入力欄ロック()
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If MsgBox("配布用ブックを作成します。元ブックは変更しませんか？", vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim copyWb As Workbook: Set copyWb = Workbooks.Add(1)
    ThisWorkbook.ActiveSheet.Copy Before:=copyWb.Worksheets(1)
    copyWb.Worksheets(1).UsedRange.Value = copyWb.Worksheets(1).UsedRange.Value
    copyWb.Worksheets(copyWb.Worksheets.Count).Delete
    copyWb.Worksheets(1).Cells.Locked = False
    copyWb.Worksheets(1).UsedRange.Locked = True
    copyWb.Worksheets(1).Protect Password:="dist"
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim path As String
    path = Application.GetSaveAsFilename("配布用ブック.xlsx", "Excelブック (*.xlsx),*.xlsx")
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If path <> "False" Then copyWb.SaveAs path, xlOpenXMLWorkbook
End Sub

