Attribute VB_Name = "modPromptHelpers"
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
