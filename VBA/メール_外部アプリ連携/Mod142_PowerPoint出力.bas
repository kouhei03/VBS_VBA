Attribute VB_Name = "Mod142_PowerPoint出力"
Option Explicit
' ファイル処理概要: PowerPointを操作を実行する。
' ファイルないコード処理詳細解説: PowerPointPDF画像出力を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、Microsoft PowerPoint、COMオートメーション利用可。
' ファイル実行例: ExcelのAlt+F8から「PowerPointPDF画像出力」を選択して実行する。

' ── PowerPointPDF画像出力 ──
Public Sub PowerPointPDF画像出力()
    Dim ppt As Object: Set ppt = CreateObject("PowerPoint.Application")
    Dim presentation As Object: Set presentation = ppt.Presentations.Open(InputBox("PPTXファイルのフルパス", "PPT出力"), , , False)
    Dim mode As String: mode = InputBox("出力形式: 1=PDF、2=PNG画像", "PPT出力", "1")
    If mode = "1" Then
        presentation.SaveAs Left$(presentation.FullName, InStrRev(presentation.FullName, ".")) & "pdf", 32
    ElseIf mode = "2" Then
        Dim folderPath As String
        With Application.FileDialog(msoFileDialogFolderPicker)
            If .Show = -1 Then folderPath = .SelectedItems(1)
        End With
        If folderPath <> "" Then presentation.Export folderPath, "PNG"
    End If
End Sub




