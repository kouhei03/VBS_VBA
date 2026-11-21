Attribute VB_Name = "Mod02_クレンジング"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod02_クレンジング
' - ファイル名: Mod02_クレンジング.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

'==================================================
' Mod02_クレンジング
' 表記揺れの一括修正
'   - 会社表記の統一（（株）→ 株式会社 など）
'   - 全角・半角の統一（英数字→半角、カナ→全角）
'   - 前後の不要なスペース削除、連続スペースの圧縮
'==================================================

'--------------------------------------------------
' メイン: 選択範囲（または使用範囲）を一括クレンジング
'--------------------------------------------------
' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub データクレンジング実行() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub データクレンジング実行()
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim target As Range
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If TypeName(Selection) = "Range" And Selection.Cells.Count > 1 Then
        Set target = Selection
    Else
        ' [ブロック解説: オブジェクト参照の設定]
        ' Workbook、Worksheet、Range、外部COMオブジェクトなどの参照を初期化します。オブジェクトが存在し、必要な参照ライブラリまたはアプリケーションが利用可能か確認します。
        Set target = ActiveSheet.UsedRange
    End If

    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If MsgBox("対象: " & target.Address(False, False) & vbCrLf & _
              "・会社表記の統一" & vbCrLf & _
              "・英数字→半角 / カナ→全角" & vbCrLf & _
              "・前後スペース削除、連続スペース圧縮" & vbCrLf & vbCrLf & _
              "実行しますか？", vbYesNo + vbQuestion) <> vbYes Then Exit Sub

    SpeedUp True
    ' [ブロック解説: エラー処理設定]
    ' エラー発生時の処理分岐を定義します。エラー番号、エラー内容、復旧処理、画面更新の復帰を確認してください。
    On Error GoTo Finally

    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim arr As Variant
    arr = target.Value
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If Not IsArray(arr) Then
        arr = CleanText(CStr(arr))
        target.Value = arr
        GoTo Finally
    End If

    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim r As Long, c As Long
    Dim changedCount As Long
    ' [ブロック解説: 繰り返し処理]
    ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
    For r = LBound(arr, 1) To UBound(arr, 1)
        ' [ブロック解説: 繰り返し処理]
        ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
        For c = LBound(arr, 2) To UBound(arr, 2)
            ' [ブロック解説: 条件分岐]
            ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
            If VarType(arr(r, c)) = vbString Then
                ' [ブロック解説: 変数・定数・設定値の宣言]
                ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
                Dim before As String, after As String
                before = arr(r, c)
                after = CleanText(before)
                ' [ブロック解説: 条件分岐]
                ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
                If before <> after Then
                    arr(r, c) = after
                    changedCount = changedCount + 1
                End If
            End If
        Next c
    Next r
    target.Value = arr

Finally:
    SpeedUp False
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If Err.Number <> 0 Then
        MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    Else
        MsgBox "クレンジング完了: " & changedCount & " セルを修正しました。", vbInformation
    End If
End Sub

'--------------------------------------------------
' 1セル分のテキストをクレンジング
'--------------------------------------------------
' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Function CleanText(ByVal s As String) As String の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Function CleanText(ByVal s As String) As String
    s = NormalizeWidth(s)
    s = NormalizeCompanyName(s)
    s = TrimSpaces(s)
    CleanText = s
End Function

'--------------------------------------------------
' 全角・半角の統一
'   英数字と記号 → 半角 / カタカナ → 全角
'--------------------------------------------------
' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Function NormalizeWidth(ByVal s As String) As String の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Function NormalizeWidth(ByVal s As String) As String
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim result As String
    Dim i As Long
    Dim ch As String

    ' [ブロック解説: 繰り返し処理]
    ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        ' [ブロック解説: 分岐条件の設定]
        ' 入力値やステータスに応じた分岐を設定します。未定義値の既定動作、すべきでない処理へのフォールバックを確認してください。
        Select Case True
            ' 全角英数字・全角スペース・主要な全角記号 → 半角
            Case ch Like "[０-９Ａ-Ｚａ-ｚ]" Or ch = "　" Or _
                 InStr("－．，＠＆（）／", ch) > 0
                result = result & StrConv(ch, vbNarrow)
            ' 半角カタカナ → 全角（濁点・半濁点は直前の文字と結合）
            Case ch Like "[ｦ-ﾟ]"
                ' [ブロック解説: 条件分岐]
                ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
                If i < Len(s) And (Mid$(s, i + 1, 1) = "ﾞ" Or Mid$(s, i + 1, 1) = "ﾟ") Then
                    result = result & StrConv(ch & Mid$(s, i + 1, 1), vbWide)
                    i = i + 1
                Else
                    result = result & StrConv(ch, vbWide)
                End If
            Case Else
                result = result & ch
        End Select
    Next i
    NormalizeWidth = result
End Function

'--------------------------------------------------
' 会社表記の統一（すべて「株式会社」等の正式表記へ）
'--------------------------------------------------
' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Function NormalizeCompanyName(ByVal s As String) As String の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Function NormalizeCompanyName(ByVal s As String) As String
    ' 半角カッコ表記（NormalizeWidth後は半角カッコになっている想定）
    s = Replace(s, "(株)", "株式会社")
    s = Replace(s, "(有)", "有限会社")
    s = Replace(s, "(合)", "合同会社")
    ' 全角カッコ表記（念のため両対応）
    s = Replace(s, "（株）", "株式会社")
    s = Replace(s, "（有）", "有限会社")
    s = Replace(s, "（合）", "合同会社")
    ' 略記
    s = Replace(s, "㈱", "株式会社")
    s = Replace(s, "㈲", "有限会社")
    NormalizeCompanyName = s
End Function

'--------------------------------------------------
' 前後スペース削除 + 文中の連続スペースを1つに圧縮
'--------------------------------------------------
' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Function TrimSpaces(ByVal s As String) As String の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Function TrimSpaces(ByVal s As String) As String
    s = Replace(s, "　", " ")           ' 全角スペース → 半角
    s = Replace(s, vbTab, " ")
    ' [ブロック解説: 条件付き繰り返し]
    ' 終了条件が満たされるまで処理を繰り返します。無限ループ防止のためカウンターと終了条件を確認してください。
    Do While InStr(s, "  ") > 0         ' 連続スペース圧縮
        s = Replace(s, "  ", " ")
    Loop
    TrimSpaces = Trim$(s)
End Function

