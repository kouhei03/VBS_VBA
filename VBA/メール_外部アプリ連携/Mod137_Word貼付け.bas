Attribute VB_Name = "Mod137_Word貼付け"
Option Explicit
' ファイル処理概要: Wordを操作を実行する。
' ファイルないコード処理詳細解説: Excel表グラフWord貼り付けを中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、Microsoft Word、COMオートメーション利用可。
' ファイル実行例: ExcelのAlt+F8から「Excel表グラフWord貼り付け」を選択して実行する。

' ── Excel表グラフWord貼り付け ──
Public Sub Excel表グラフWord貼り付け()
    If Selection Is Nothing Then Exit Sub
    Selection.Copy
    Dim word As Object: Set word = CreateObject("Word.Application")
    word.Visible = True
    Dim document As Object: Set document = word.Documents.Add
    document.Content.Paste
    MsgBox "選択範囲を貼り付けました。", vbInformation
End Sub




