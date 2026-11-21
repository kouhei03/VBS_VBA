Attribute VB_Name = "Mod119_ブックリンク更新"
Option Explicit

Public Sub 複数ブックリンク更新制御()
    Dim mode As String
    mode = InputBox("リンク更新を制御: 1=更新する、2=更新しない、3=状態確認", "リンク更新", "2")
    If mode = "1" Then
        Application.AskToUpdateLinks = True
        ThisWorkbook.UpdateLinks = xlUpdateLinksAlways
    ElseIf mode = "2" Then
        Application.AskToUpdateLinks = False
        ThisWorkbook.UpdateLinks = xlUpdateLinksNever
    ElseIf mode = "3" Then
        Dim links As Variant
        links = ThisWorkbook.LinkSources(xlExcelLinks)
        If IsEmpty(links) Then
            MsgBox "外部リンクはありません。", vbInformation
        Else
            MsgBox "ブックの外部リンク数: " & UBound(links) - LBound(links) + 1, vbInformation
        End If
    End If
End Sub
