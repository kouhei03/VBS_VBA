Attribute VB_Name = "Mod103_ブック形式変換"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: ブック形式変換保存を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「ブック形式変換保存」を選択して実行する。

' ── ブック形式変換保存 ──
Public Sub ブック形式変換保存()
    Dim mode As String
    mode = InputBox("保存形式を入力: 1=xlsx、2=xlsm、3=xls、4=xlsb", "形式変換", "1")
    Dim destination As Long
    Select Case mode
        Case "1": destination = xlOpenXMLWorkbook
        Case "2": destination = xlOpenXMLWorkbookMacroEnabled
        Case "3": destination = xlExcel8
        Case "4": destination = xlExcel12
        Case Else: Exit Sub
    End Select
    Dim baseName As String
    baseName = Left$(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1)
    ThisWorkbook.SaveAs baseName & "_" & Format$(Now, "yyyymmdd") & GetExtension(mode), destination
End Sub

' ── GetExtension ──
Private Function GetExtension(ByVal mode As String) As String
    Select Case mode
        Case "1": GetExtension = ".xlsx"
        Case "2": GetExtension = ".xlsm"
        Case "3": GetExtension = ".xls"
        Case "4": GetExtension = ".xlsb"
    End Select
End Function




