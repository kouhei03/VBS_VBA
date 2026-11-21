Attribute VB_Name = "Mod184_ヘッダーフッター"
Option Explicit

Public Sub ヘッダーフッターページ番号日時社名設定()
    Dim companyName As String: companyName = InputBox("社名", "ヘッダー/フッター")
    With ActiveSheet.PageSetup
        .LeftHeader = companyName
        .CenterHeader = "&F"
        .RightHeader = "&D &T"
        .LeftFooter = "社名: " & companyName
        .CenterFooter = "Page &P / &N"
    End With
End Sub
