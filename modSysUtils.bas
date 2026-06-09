Attribute VB_Name = "modSysUtils"
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
        (ByVal hwnd As LongPtr, ByVal lpOperation As String, ByVal lpFile As String, _
        ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As LongPtr
#Else
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
        (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
        ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
#End If

' Copies text to the Windows Clipboard using late-bound MSForms.DataObject
Public Sub CopyToClipboard(ByVal TextToCopy As String)
    Dim DataObj As Object
    On Error Resume Next
    
    ' Using the CLSID for MSForms.DataObject avoids needing the FM20.dll reference explicitly
    Set DataObj = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    If DataObj Is Nothing Then
        ' Fallback using HTMLFile if DataObject fails (sometimes happens on 64-bit Office)
        FallbackCopyToClipboard TextToCopy
        Exit Sub
    End If
    
    DataObj.SetText TextToCopy
    DataObj.PutInClipboard
    Set DataObj = Nothing
    On Error GoTo 0
End Sub

Private Sub FallbackCopyToClipboard(ByVal TextToCopy As String)
    Dim html As Object
    On Error Resume Next
    Set html = CreateObject("htmlfile")
    html.ParentWindow.ClipboardData.SetData "text", TextToCopy
    Set html = Nothing
    On Error GoTo 0
End Sub

' Opens a specified URL in the system's default web browser
Public Sub OpenURL(ByVal url As String)
    On Error Resume Next
    ShellExecute 0, "open", url, vbNullString, vbNullString, 1
    On Error GoTo 0
End Sub

' Content fingerprint for transport attestation (MKS TSA Profile s4.1 / s9.1).
' This is a deterministic, dependency-free integrity fingerprint -- two 31-bit
' polynomial-hash lanes rendered as 16 hex chars -- NOT a cryptographic hash.
' Its job is the clipboard failure mode the Profile names: detecting that a
' payload was "pasted unmodified" (accidental corruption / truncation /
' mis-paste), not resisting an adversary. We avoid CAPICOM/.NET COM so it runs
' on locked-down Office with no references, matching the rest of this tool.
'
' Arithmetic is done entirely in Double, which represents every integer below
' 2^53 exactly. With primes < 2^31 and small multipliers, h*B + c stays under
' 2^38, so there is no precision loss and no Long-overflow on bitwise ops
' (there are none). Lane 2 folds character position in, so a reordering that
' preserves the multiset of characters still changes the digest.
Public Function ArContentFingerprint(ByVal s As String) As String
    Dim i As Long
    Dim ch As Double
    Dim h1 As Double, h2 As Double

    Const B1 As Double = 131#
    Const P1 As Double = 2147483647#   ' 2^31 - 1 (prime)
    Const B2 As Double = 137#
    Const P2 As Double = 2147483629#   ' largest prime below 2^31 - 18

    h1 = 2166136261# - P1   ' fold the classic FNV seed into range
    h2 = 1099511628#

    For i = 1 To Len(s)
        ch = AscW(Mid$(s, i, 1)) And &HFFFF&

        h1 = h1 * B1 + ch + 1#
        h1 = h1 - Int(h1 / P1) * P1

        h2 = h2 * B2 + ch + (i Mod 251) + 1#
        h2 = h2 - Int(h2 / P2) * P2
    Next i

    ArContentFingerprint = Hex31(h1) & Hex31(h2)
End Function

' Render a Double holding a value in [0, 2^31) as 8 hex chars.
Private Function Hex31(ByVal v As Double) As String
    Dim s As String
    s = Hex$(CLng(v))
    If Len(s) < 8 Then s = String$(8 - Len(s), "0") & s
    Hex31 = Right$(s, 8)
End Function
