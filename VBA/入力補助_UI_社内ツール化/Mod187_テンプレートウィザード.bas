Attribute VB_Name = "Mod187_テンプレートウィザード"
Option Explicit
' ファイル処理概要: ブック・シートを処理を実行する。
' ファイルないコード処理詳細解説: テンプレート選択新規作成ウィザードを中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「テンプレート選択新規作成ウィザード」を選択して実行する。

' ── テンプレート選択新規作成ウィザード ──
Public Sub テンプレート選択新規作成ウィザード()
    Dim template As String
    template = InputBox("テンプレートを選択: 1=請求、2=在庫、3=申請", "テンプレートウィザード")
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets.Add
    Select Case template
        Case "1": ws.Name = "請求テンプレート"
        Case "2": ws.Name = "在庫テンプレート"
        Case "3": ws.Name = "申請テンプレート"
        Case Else: ws.Delete: Exit Sub
    End Select
    ws.Range("A1").Value = "項目"
    ws.Range("B1").Value = "入力値"
End Sub




