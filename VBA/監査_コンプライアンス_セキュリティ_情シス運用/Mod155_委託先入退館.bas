Attribute VB_Name = "Mod155_委託先入退館"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 委託先ベンダー入退館アクセス権限管理を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「委託先ベンダー入退館アクセス権限管理」を選択して実行する。

' ── 委託先ベンダー入退館アクセス権限管理 ──
Public Sub 委託先ベンダー入退館アクセス権限管理()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim r As Long
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If CStr(ws.Cells(r, 4).Value) = "退館済" Or CStr(ws.Cells(r, 5).Value) = "無効" Then
            ws.Cells(r, 6).Value = "対応済"
        Else
            ws.Cells(r, 6).Value = "要確認"
        End If
    Next r
End Sub




