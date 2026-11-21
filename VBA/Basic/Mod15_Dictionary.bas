Attribute VB_Name = "Mod15_Dictionary"
Option Explicit
' ファイル処理概要: 辞書でキー照合を実行する。
' ファイルないコード処理詳細解説: Dictionaryで件数集計を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「Dictionaryで件数集計」を選択して実行する。

' ── Dictionaryで件数集計 ──
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




