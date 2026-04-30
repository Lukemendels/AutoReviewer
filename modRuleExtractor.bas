Attribute VB_Name = "modRuleExtractor"
Option Explicit

Public Sub ExtractRulesAndPrompt()
    Dim wdApp As Object
    Dim wdDoc As Object
    Dim rev As Object
    Dim ws As Object
    Dim fso As Object
    Dim ts As Object
    Dim exportPath As String
    Dim exportText As String
    Dim revCount As Long
    Dim promptText As String
    Dim gptUrl As String
    
    ' Connect to active Word instance
    On Error Resume Next
    Set wdApp = GetObject(, "Word.Application")
    If wdApp Is Nothing Then
        MsgBox "Word is not running or no active document found.", vbExclamation
        Exit Sub
    End If
    
    Set wdDoc = wdApp.ActiveDocument
    If wdDoc Is Nothing Then
        MsgBox "No active Word document found.", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Check for revisions
    If wdDoc.Revisions.Count = 0 Then
        MsgBox "No track changes found in the active document. Please open a document with your boss's tracked changes.", vbInformation
        Exit Sub
    End If
    
    ' Setup file system and export path
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ws = CreateObject("WScript.Shell")
    exportPath = ws.SpecialFolders("MyDocuments") & "\ExtractedRevisions.md"
    
    ' Overwrite file if it exists, use Unicode (UTF-16) for Word text
    Set ts = fso.CreateTextFile(exportPath, True, True)
    
    ' Build Markdown content
    ts.WriteLine "# Track Changes Extraction"
    ts.WriteLine "Extracted from: " & wdDoc.Name
    ts.WriteLine "Total Revisions: " & wdDoc.Revisions.Count
    ts.WriteLine ""
    
    revCount = 0
    For Each rev In wdDoc.Revisions
        ' Focus on Insertions (1) and Deletions (2)
        If rev.Type = 1 Or rev.Type = 2 Then
            revCount = revCount + 1
            ts.WriteLine "## Revision " & revCount
            
            Dim revTypeStr As String
            If rev.Type = 1 Then
                revTypeStr = "Insertion"
            Else
                revTypeStr = "Deletion"
            End If
            
            ts.WriteLine "**Type:** " & revTypeStr
            ts.WriteLine "**Author:** " & rev.Author
            
            ' Try to extract context (the sentence or paragraph)
            Dim contextText As String
            On Error Resume Next
            contextText = rev.Range.Paragraphs(1).Range.Text
            ' Clean up context text
            contextText = Replace(contextText, vbCr, "")
            contextText = Replace(contextText, vbLf, "")
            On Error GoTo 0
            
            ts.WriteLine "**Context:** " & contextText
            
            Dim cleanRevText As String
            cleanRevText = Replace(rev.Range.Text, vbCr, "")
            cleanRevText = Replace(cleanRevText, vbLf, "")
            ts.WriteLine "**Revision Text:** " & cleanRevText
            ts.WriteLine ""
        End If
    Next rev
    
    ts.Close
    
    ' Prepare Prompt
    promptText = "I am providing a markdown file containing Track Changes extracted from my boss's document. " & _
                 "Please analyze these before-and-after edits, deduce the implicit style rules, tone preferences, and formatting habits. " & _
                 "Output a comprehensive SKILL.md file that can be used to instruct you (or another LLM) to simulate this exact review style in the future."
    
    ' Copy to clipboard
    modPromptHelpers.CopyToClipboard promptText
    
    ' Read Custom GPT URL from config or fallback
    gptUrl = "https://chatgpt.com/"
    On Error Resume Next
    Dim configUrl As String
    ' Requires ConfigHelpers.bas to be present
    configUrl = ConfigHelpers.GetConfigValue("CustomGptUrl")
    If Len(configUrl) > 0 Then gptUrl = configUrl
    On Error GoTo 0
    
    ' Launch URL
    modPromptHelpers.OpenURL gptUrl
    
    MsgBox "Extracted " & revCount & " revisions to:" & vbCrLf & exportPath & vbCrLf & vbCrLf & _
           "The prompt has been copied to your clipboard!" & vbCrLf & _
           "1. The browser has been opened to your GPT." & vbCrLf & _
           "2. Paste the prompt (Ctrl+V)." & vbCrLf & _
           "3. Upload/Drop the ExtractedRevisions.md file.", vbInformation, "Extraction Complete"

End Sub
