Attribute VB_Name = "Mod09_エラー番号分岐"
Option Explicit
' ファイル処理概要: エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: エラー番号別処理を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「エラー番号別処理」を選択して実行する。

' ── エラー番号別処理 ──
Public Sub エラー番号別処理()
    On Error GoTo Failed

    ' ここに処理を記述
    Exit Sub
Failed:
    Select Case Err.Number
        Case 9
            MsgBox "指定されたシートまたは配列要素が存在しません。", vbExclamation
        Case 13
            MsgBox "型変換エラーです。入力値を確認してください。", vbExclamation
        Case 1004
            MsgBox "Excelの範囲操作でエラーが発生しました。", vbExclamation
        Case Else
            MsgBox "予期しないエラー " & Err.Number & ": " & Err.Description, vbCritical
    End Select
End Sub




