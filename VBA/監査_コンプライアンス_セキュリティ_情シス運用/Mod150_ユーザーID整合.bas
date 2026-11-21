Attribute VB_Name = "Mod150_ユーザーID整合"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: ユーザーID命名規則メールエイリアス所属整合を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「ユーザーID命名規則メールエイリアス所属整合」を選択して実行する。

' ── ユーザーID命名規則メールエイリアス所属整合 ──
Public Sub ユーザーID命名規則メールエイリアス所属整合()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim prefix As String: prefix = InputBox("許容するユーザーID接頭辞", "ID整合", "USR")
    Dim r As Long
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If Left$(CStr(ws.Cells(r, 1).Value), Len(prefix)) <> prefix _
           Or InStr(CStr(ws.Cells(r, 2).Value), "@") = 0 Then
            ws.Rows(r).Interior.Color = RGB(255, 199, 206)
        End If
    Next r
End Sub




