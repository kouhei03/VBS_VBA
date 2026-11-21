Attribute VB_Name = "Mod15_Dictionary"
Option Explicit

Public Sub Dictionaryで件数集計(ByVal rng As Range)
    Dim counts As Object
    Set counts = CreateObject("Scripting.Dictionary")
    counts.CompareMode = vbTextCompare

    Dim value As Variant
    For Each value In rng.Value
        If Len(Trim$(CStr(value))) > 0 Then
            If counts.Exists(CStr(value)) Then
                counts(CStr(value)) = counts(CStr(value)) + 1
            Else
                counts.Add CStr(value), 1
            End If
        End If
    Next value

    Dim key As Variant
    For Each key In counts.Keys
        Debug.Print key & ": " & counts(key)
    Next key
End Sub
