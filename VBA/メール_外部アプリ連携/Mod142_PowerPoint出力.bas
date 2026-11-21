Attribute VB_Name = "Mod142_PowerPoint出力"
Option Explicit

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
