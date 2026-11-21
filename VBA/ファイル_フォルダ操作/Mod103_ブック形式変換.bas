Attribute VB_Name = "Mod103_ブック形式変換"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod103_ブック形式変換
' - ファイル名: Mod103_ブック形式変換.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub ブック形式変換保存() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub ブック形式変換保存()
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim mode As String
    mode = InputBox("保存形式を入力: 1=xlsx、2=xlsm、3=xls、4=xlsb", "形式変換", "1")
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim destination As Long
    ' [ブロック解説: 分岐条件の設定]
    ' 入力値やステータスに応じた分岐を設定します。未定義値の既定動作、すべきでない処理へのフォールバックを確認してください。
    Select Case mode
        Case "1": destination = xlOpenXMLWorkbook
        Case "2": destination = xlOpenXMLWorkbookMacroEnabled
        Case "3": destination = xlExcel8
        Case "4": destination = xlExcel12
        Case Else: Exit Sub
    End Select
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim baseName As String
    baseName = Left$(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1)
    ThisWorkbook.SaveAs baseName & "_" & Format$(Now, "yyyymmdd") & GetExtension(mode), destination
End Sub

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Private Function GetExtension(ByVal mode As String) As String の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Private Function GetExtension(ByVal mode As String) As String
    ' [ブロック解説: 分岐条件の設定]
    ' 入力値やステータスに応じた分岐を設定します。未定義値の既定動作、すべきでない処理へのフォールバックを確認してください。
    Select Case mode
        Case "1": GetExtension = ".xlsx"
        Case "2": GetExtension = ".xlsm"
        Case "3": GetExtension = ".xls"
        Case "4": GetExtension = ".xlsb"
    End Select
End Function

