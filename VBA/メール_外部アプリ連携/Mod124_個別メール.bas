Attribute VB_Name = "Mod124_個別メール"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 担当者取引先別個別メール生成を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、クラシックOutlook、Outlookプロファイル。
' ファイル実行例: ExcelのAlt+F8から「担当者取引先別個別メール生成」を選択して実行する。

' ── 担当者取引先別個別メール生成 ──
Public Sub 担当者取引先別個別メール生成()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim outlook As Object: Set outlook = GetOutlook()
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim r As Long
    For r = 2 To lastRow
        If Trim$(CStr(ws.Cells(r, 1).Value)) <> "" Then
            Dim mail As Object: Set mail = outlook.CreateItem(0)
            mail.To = CStr(ws.Cells(r, 2).Value)
            mail.Subject = CStr(ws.Cells(r, 3).Value)
            mail.Body = CStr(ws.Cells(r, 1).Value) & "様（" & CStr(ws.Cells(r, 4).Value) & "）" & vbCrLf & CStr(ws.Cells(r, 5).Value)
            mail.Save
        End If
    Next r
    MsgBox "個別メールの下書きを生成しました。", vbInformation
End Sub




