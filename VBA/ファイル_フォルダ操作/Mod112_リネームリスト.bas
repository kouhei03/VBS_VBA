Attribute VB_Name = "Mod112_リネームリスト"
Option Explicit

Public Sub 処理対象リストに基づき一括リネーム()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "A列に変更前パス、B列に変更後パスを設定してください。", vbExclamation
        Exit Sub
    End If
    If MsgBox(lastRow - 1 & " 件をリネームしますか？", vbYesNo + vbQuestion) <> vbYes Then Exit Sub

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim r As Long
    For r = 2 To lastRow
        Dim sourcePath As String: sourcePath = CStr(ws.Cells(r, 1).Value)
        Dim destinationPath As String: destinationPath = CStr(ws.Cells(r, 2).Value)
        If fso.FileExists(sourcePath) And destinationPath <> "" Then
            If fso.FileExists(destinationPath) Then
                ws.Cells(r, 3).Value = "移動先が既に存在"
            Else
                fso.MoveFile sourcePath, destinationPath
                ws.Cells(r, 3).Value = "完了"
            End If
        Else
            ws.Cells(r, 3).Value = "元ファイルなし"
        End If
    Next r
End Sub
