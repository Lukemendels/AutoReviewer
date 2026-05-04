Attribute VB_Name = "modTrainingOrchestrator"
Option Explicit

' Pass 1: Cluster
Public Sub RunReducePass1()
    Dim activePersona As String
    Dim corpusPath As String
    Dim prompt As String
    Dim fso As Object
    
    activePersona = ConfigHelpers.GetConfigValue("ActivePersona")
    If activePersona = "" Then
        MsgBox "Please select an Active Persona first.", vbExclamation
        Exit Sub
    End If
    
    corpusPath = ThisWorkbook.path & "\" & activePersona & "_corpus.jsonl"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(corpusPath) Then
        MsgBox "Corpus file not found: " & corpusPath, vbExclamation
        Exit Sub
    End If
    
    prompt = "Please review the attached corpus of revisions and comments. Cluster these into pattern categories based on the reviewer's implicit style preferences."
    
    modPromptHelpers.CopyToClipboard prompt
    
    MsgBox "Reduce Pass 1 Prompt copied to clipboard." & vbCrLf & vbCrLf & _
           "1. Open a new chat in DHSChat." & vbCrLf & _
           "2. Drag and drop the corpus file from the Explorer window that will open." & vbCrLf & _
           "3. Paste the prompt and send.", vbInformation
           
    ' Open Explorer and select the file
    Shell "explorer.exe /select,""" & corpusPath & """", vbNormalFocus
    
    ' Launch DHSChat or default browser to chatgpt
    Dim url As String
    url = ConfigHelpers.GetConfigValue("CustomGptUrl", "https://chatgpt.com/")
    If InStr(LCase(url), "http") > 0 Then
        Shell "explorer.exe """ & url & """", vbNormalFocus
    End If
End Sub

' Pass 2: Extract Heuristics
Public Sub RunReducePass2()
    Dim prompt As String
    prompt = "Based on the clusters you just identified, extract a clear, actionable heuristic for each category. Explain the rationale behind each heuristic."
    
    modPromptHelpers.CopyToClipboard prompt
    
    MsgBox "Reduce Pass 2 Prompt copied to clipboard." & vbCrLf & vbCrLf & _
           "1. Paste this prompt into the same DHSChat conversation and send.", vbInformation
End Sub

' Pass 3: Synthesize SKILL.md
Public Sub RunReducePass3()
    Dim prompt As String
    prompt = "Synthesize the extracted heuristics into a comprehensive SKILL.md file for a DHSChat Assistant. " & _
             "Include the style guidance and the strict JSONL output contract required for bookmark-targeted edits. " & _
             "Return ONLY the markdown code block containing the SKILL.md."
    
    modPromptHelpers.CopyToClipboard prompt
    
    MsgBox "Reduce Pass 3 Prompt copied to clipboard." & vbCrLf & vbCrLf & _
           "1. Paste this prompt into the same DHSChat conversation and send." & vbCrLf & _
           "2. Once generated, save the SKILL.md and run Save SKILL.md.", vbInformation
End Sub

' Save SKILL.md
Public Sub SaveSkillMd()
    Dim activePersona As String
    Dim skillPath As String
    Dim fso As Object
    Dim ts As Object
    Dim skillContent As String
    Dim dataObj As Object
    
    activePersona = ConfigHelpers.GetConfigValue("ActivePersona")
    If activePersona = "" Then
        MsgBox "Please select an Active Persona first.", vbExclamation
        Exit Sub
    End If
    
    skillPath = ThisWorkbook.path & "\" & activePersona & "_SKILL.md"
    
    On Error Resume Next
    Set dataObj = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}") ' MSForms.DataObject
    dataObj.GetFromClipboard
    skillContent = dataObj.GetText
    On Error GoTo 0
    
    If skillContent = "" Then
        MsgBox "Please copy the generated SKILL.md text to the clipboard first.", vbExclamation
        Exit Sub
    End If
    
    ' Strip markdown formatting if present
    If Left(skillContent, 3) = "```" Then
        ' Simple cleanup
        skillContent = Replace(skillContent, "```markdown", "")
        skillContent = Replace(skillContent, "```", "")
        skillContent = Trim(skillContent)
    End If
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(skillPath, 2, True, -1) ' 2=Write, True=Create, -1=Unicode
    ts.Write skillContent
    ts.Close
    
    modPersonaRegistry.UpsertPersona activePersona, skillMdPath:=skillPath
    
    MsgBox "SKILL.md saved to: " & skillPath & vbCrLf & vbCrLf & _
           "Next Step: Create a new DHSChat Assistant, paste the SKILL.md into its system prompt, " & _
           "and save the Assistant URL into the Persona Registry.", vbInformation
End Sub
