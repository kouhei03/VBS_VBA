Attribute VB_Name = "Mod162_初期値"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: フォーム初期値現在日付ユーザー情報セットを中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「フォーム初期値現在日付ユーザー情報セット」を選択して実行する。

' ── フォーム初期値現在日付ユーザー情報セット ──
Public Sub フォーム初期値現在日付ユーザー情報セット()
    Dim ws As Worksheet: Set ws = ActiveSheet
    ws.Range("A1").Value = "登録日"
    ws.Range("B1").Value = Date
    ws.Range("A2").Value = "ログインユーザー"
    ws.Range("B2").Value = CreateObject("WScript.Network").UserName
End Sub




