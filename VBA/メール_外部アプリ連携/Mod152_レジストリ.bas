Attribute VB_Name = "Mod152_レジストリ"
Option Explicit
' ファイル処理概要: Windowsシェルを操作を実行する。
' ファイルないコード処理詳細解説: レジストリ値読み書きを中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、Windowsファイルシステム、必要な実行権限。
' ファイル実行例: ExcelのAlt+F8から「レジストリ値読み書き」を選択して実行する。

' ── レジストリ値読み書き ──
Public Sub レジストリ値読み書き()
    Dim shell As Object: Set shell = CreateObject("WScript.Shell")
    Dim mode As String: mode = InputBox("操作を入力: 1=読み取り、2=書き込み", "レジストリ")
    Dim keyPath As String: keyPath = InputBox("レジストリパスを入力", "レジストリ")
    If mode = "1" Then
        On Error Resume Next
        MsgBox CStr(shell.RegRead(keyPath)), vbInformation
        If Err.Number <> 0 Then MsgBox "値を読み取れません。", vbCritical
    ElseIf mode = "2" Then
        Dim value As String: value = InputBox("保存する文字列値", "レジストリ")
        shell.RegWrite keyPath, value, "REG_SZ"
        MsgBox "レジストリへ保存しました。", vbInformation
    End If
End Sub




