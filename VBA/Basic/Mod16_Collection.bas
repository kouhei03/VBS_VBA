Attribute VB_Name = "Mod16_Collection"
Option Explicit

Public Sub Collectionで可変リスト管理()
    Dim items As New Collection
    items.Add "りんご"
    items.Add "みかん"
    items.Add "ぶどう"

    Dim i As Long
    For i = 1 To items.Count
        Debug.Print items(i)
    Next i

    items.Remove 2
End Sub
