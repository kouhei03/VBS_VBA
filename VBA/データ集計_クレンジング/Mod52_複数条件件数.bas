Attribute VB_Name = "Mod52_複数条件件数"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 複数条件で件数カウントを中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「複数条件で件数カウント」を選択して実行する。

' ── 複数条件で件数カウント ──
Public Sub 複数条件で件数カウント()
    Dim col1 As Long: col1 = CLng(InputBox("条件列1", "件数カウント", "1"))
    Dim value1 As String: value1 = InputBox("条件値1", "件数カウント")
    Dim col2 As Long: col2 = CLng(InputBox("条件列2", "件数カウント", "2"))
    Dim value2 As String: value2 = InputBox("条件値2", "件数カウント")
    Dim count As Long
    Dim r As Long
    For r = 2 To GetLastRow(ActiveSheet)
        If CStr(ActiveSheet.Cells(r, col1).Value) = value1 And _
           CStr(ActiveSheet.Cells(r, col2).Value) = value2 Then count = count + 1
    Next r
    MsgBox "該当件数: " & count & " 件", vbInformation
End Sub




