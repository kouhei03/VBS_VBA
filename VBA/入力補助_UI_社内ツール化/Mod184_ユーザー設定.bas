Attribute VB_Name = "Mod184_ユーザー設定"
Option Explicit
' ファイル処理概要: Windowsシェルを操作、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: ユーザー部署別設定保存復元を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、Windowsファイルシステム、必要な実行権限。
' ファイル実行例: ExcelのAlt+F8から「ユーザー部署別設定保存復元」を選択して実行する。

' ── ユーザー部署別設定保存復元 ──
Public Sub ユーザー部署別設定保存復元()
    Dim key As String: key = CreateObject("WScript.Network").UserName
    Dim setting As String: setting = InputBox("保存または復元する設定値", "ユーザー設定")
    If setting = "" Then
        Dim restored As String
        On Error Resume Next
        restored = CStr(CreateObject("WScript.Shell").RegRead("HKCU\Software\VBAInternalTools\" & key))
        On Error GoTo 0
        MsgBox restored, vbInformation
    Else
        CreateObject("WScript.Shell").RegWrite "HKCU\Software\VBAInternalTools\" & key, setting, "REG_SZ"
        MsgBox "設定を保存しました。", vbInformation
    End If
End Sub




