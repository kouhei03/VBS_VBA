Attribute VB_Name = "Mod150_クリップボード"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: クリップボードテキストコピー取得を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「クリップボードテキストコピー取得」を選択して実行する。

' ── クリップボードテキストコピー取得 ──
Public Sub クリップボードテキストコピー取得()
    Dim mode As String: mode = InputBox("操作を入力: 1=コピー、2=取得", "クリップボード")
    Dim data As Object: Set data = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    If mode = "1" Then
        data.SetText InputBox("コピーするテキスト", "クリップボード")
        data.PutInClipboard
    ElseIf mode = "2" Then
        data.GetFromClipboard
        MsgBox data.GetText, vbInformation
    End If
End Sub




