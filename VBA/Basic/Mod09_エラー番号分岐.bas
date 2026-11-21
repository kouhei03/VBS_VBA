Attribute VB_Name = "Mod09_エラー番号分岐"
Option Explicit

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
