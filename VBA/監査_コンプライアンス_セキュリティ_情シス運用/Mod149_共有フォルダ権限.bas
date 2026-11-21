Attribute VB_Name = "Mod149_共有フォルダ権限"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 共有フォルダドライブ権限外部公開広範囲共有抽出を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「共有フォルダドライブ権限外部公開広範囲共有抽出」を選択して実行する。

' ── 共有フォルダドライブ権限外部公開広範囲共有抽出 ──
Public Sub 共有フォルダドライブ権限外部公開広範囲共有抽出()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim r As Long
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        Dim audience As String: audience = CStr(ws.Cells(r, 3).Value)
        If audience = "Everyone" Or audience = "全社" Or audience = "外部" Then
            ws.Rows(r).Interior.Color = RGB(255, 199, 206)
        End If
    Next r
End Sub




