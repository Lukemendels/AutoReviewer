Attribute VB_Name = "modTrainingPipeline"
Option Explicit


' Takes an open Word Document object and prompts user to pick a target author.
Public Function SelectTargetAuthor(ByVal wdDoc As Object) As String
    Dim rev As Object
    Dim cmt As Object
    Dim authors As Object
    Dim authorName As Variant
    Dim msg As String
    Dim i As Integer
    Dim authorList() As String
    Dim choice As String
    Dim choiceNum As Integer
    
    Set authors = CreateObject("Scripting.Dictionary")
    
    ' Tally revisions
    For Each rev In wdDoc.Revisions
        authorName = rev.Author
        If Not authors.exists(authorName) Then
            authors.Add authorName, 1
        Else
            authors(authorName) = authors(authorName) + 1
        End If
    Next rev
    
    ' Tally comments
    For Each cmt In wdDoc.Comments
        authorName = cmt.Author
        If Not authors.exists(authorName) Then
            authors.Add authorName, 1
        Else
            authors(authorName) = authors(authorName) + 1
        End If
    Next cmt
    
    If authors.Count = 0 Then
        MsgBox "No revisions or comments found in this document.", vbExclamation
        SelectTargetAuthor = ""
        Exit Function
    End If
    
    ReDim authorList(1 To authors.Count)
    msg = "Select the target author by entering their number:" & vbCrLf & vbCrLf
    
    i = 1
    For Each authorName In authors.Keys
        authorList(i) = CStr(authorName)
        msg = msg & i & ". " & authorName & " (" & authors(authorName) & " edits/comments)" & vbCrLf
        i = i + 1
    Next authorName
    
    choice = InputBox(msg, "Select Target Author")
    
    If Trim(choice) = "" Then
        SelectTargetAuthor = ""
        Exit Function
    End If
    
    If IsNumeric(choice) Then
        choiceNum = CInt(choice)
        If choiceNum >= 1 And choiceNum <= authors.Count Then
            SelectTargetAuthor = authorList(choiceNum)
        Else
            MsgBox "Invalid selection.", vbExclamation
            SelectTargetAuthor = ""
        End If
    Else
        MsgBox "Please enter a number.", vbExclamation
        SelectTargetAuthor = ""
    End If
End Function


Public Sub AddDocToCorpus()
    Dim wdApp As Object
    Dim wdDoc As Object
    Dim docPath As String
    Dim targetAuthor As String
    Dim fd As FileDialog
    
    Dim activePersona As String
    activePersona = modAppCore.GetConfigValue("ActivePersona")
    
    If activePersona = "" Then
        MsgBox "Please select an Active Persona first.", vbExclamation
        Exit Sub
    End If
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "Select Word Document for Training Corpus"
    fd.Filters.Clear
    fd.Filters.Add "Word Documents", "*.docx"
    If fd.Show <> -1 Then Exit Sub
    docPath = fd.SelectedItems(1)
    
    On Error Resume Next
    Set wdApp = GetObject(, "Word.Application")
    If wdApp Is Nothing Then
        Set wdApp = CreateObject("Word.Application")
    End If
    On Error GoTo 0
    
    wdApp.Visible = True
    Set wdDoc = wdApp.Documents.Open(docPath)
    
    targetAuthor = modTrainingPipeline.SelectTargetAuthor(wdDoc)
    
    If targetAuthor = "" Then
        wdDoc.Close False
        Exit Sub
    End If
    
    ' Step: Compute target date range
    Dim minDate As Date
    Dim firstTargetFound As Boolean
    Dim rev As Object
    Dim cmt As Object
    
    firstTargetFound = False
    
    For Each rev In wdDoc.Revisions
        If StrComp(rev.Author, targetAuthor, vbTextCompare) = 0 Then
            If Not firstTargetFound Then
                minDate = rev.Date
                firstTargetFound = True
            Else
                If rev.Date < minDate Then minDate = rev.Date
            End If
        End If
    Next rev
    
    For Each cmt In wdDoc.Comments
        If StrComp(cmt.Author, targetAuthor, vbTextCompare) = 0 Then
            If Not firstTargetFound Then
                minDate = cmt.Date
                firstTargetFound = True
            Else
                If cmt.Date < minDate Then minDate = cmt.Date
            End If
        End If
    Next cmt
    
    ' Build baseline: Accept non-target revisions before target earliest
    Dim i As Long
    For i = wdDoc.Revisions.Count To 1 Step -1
        Set rev = wdDoc.Revisions(i)
        If StrComp(rev.Author, targetAuthor, vbTextCompare) <> 0 Then
            If firstTargetFound And rev.Date < minDate Then
                rev.Accept
            End If
        End If
    Next i
    
    ' Stamp bookmarks
    modWordUtils.StampDocWithArBookmarks wdDoc
    
    ' Extract records to JSONL
    Dim fso As Object
    Dim ts As Object
    Dim corpusPath As String
    
    corpusPath = ThisWorkbook.path & "\" & activePersona & "_corpus.jsonl"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(corpusPath, 8, True, -1) ' 8=Append, True=Create, -1=Unicode
    
    ' Revisions output
    Dim jsonLine As String
    For Each rev In wdDoc.Revisions
        If StrComp(rev.Author, targetAuthor, vbTextCompare) = 0 Then
            ' Simplified JSON construction for V1
            jsonLine = "{""doc_id"":""" & Replace(wdDoc.name, "\", "\\") & """" & _
                       ",""target_author"":""" & Replace(targetAuthor, """", "\""") & """" & _
                       ",""record_type"":""revision""" & _
                       ",""date"":""" & rev.Date & """}"
            ts.WriteLine jsonLine
        End If
    Next rev
    
    ' Comments output
    For Each cmt In wdDoc.Comments
        If StrComp(cmt.Author, targetAuthor, vbTextCompare) = 0 Then
            jsonLine = "{""doc_id"":""" & Replace(wdDoc.name, "\", "\\") & """" & _
                       ",""target_author"":""" & Replace(targetAuthor, """", "\""") & """" & _
                       ",""record_type"":""comment""" & _
                       ",""date"":""" & cmt.Date & """}"
            ts.WriteLine jsonLine
        End If
    Next cmt
    
    ts.Close
    
    ' Update Registry
    modAppCore.UpsertPersona activePersona, corpusPath:=corpusPath, incrementTrainingCount:=True
    
    wdDoc.Close False
    Set wdDoc = Nothing
    Set wdApp = Nothing
    DoEvents
    
    MsgBox "Document added to corpus successfully.", vbInformation
End Sub


' Pass 1: Cluster
Public Sub RunReducePass1()
    Dim activePersona As String
    Dim corpusPath As String
    Dim prompt As String
    Dim fso As Object
    
    activePersona = modAppCore.GetConfigValue("ActivePersona")
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
    
    modSysUtils.CopyToClipboard prompt
    
    MsgBox "Reduce Pass 1 Prompt copied to clipboard." & vbCrLf & vbCrLf & _
           "1. Open a new chat in DHSChat." & vbCrLf & _
           "2. Drag and drop the corpus file from the Explorer window that will open." & vbCrLf & _
           "3. Paste the prompt and send.", vbInformation
           
    ' Open Explorer and select the file
    Shell "explorer.exe /select,""" & corpusPath & """", vbNormalFocus
    
    ' Launch DHSChat or default browser to chat
    
    url = modAppCore.GetConfigValue("CustomGptUrl", "https://chat.dhs.gov/workspaces/4cf75bdf-de55-4f01-8c3f-0444ace52010")
    If InStr(LCase(url), "http") > 0 Then
        Shell "explorer.exe """ & url & """", vbNormalFocus
    End If
End Sub

' Pass 2: Extract Heuristics
Public Sub RunReducePass2()
    Dim prompt As String
    prompt = "Based on the clusters you just identified, extract a clear, actionable heuristic for each category. Explain the rationale behind each heuristic."
    
    modSysUtils.CopyToClipboard prompt
    
    MsgBox "Reduce Pass 2 Prompt copied to clipboard." & vbCrLf & vbCrLf & _
           "1. Paste this prompt into the same DHSChat conversation and send.", vbInformation
End Sub

' Pass 3: Synthesize SKILL.md
Public Sub RunReducePass3()
    Dim prompt As String
    prompt = "Synthesize the extracted heuristics into a SKILL.md for a HOT co-thinker " & _
             "DHSChat assistant (the reviewer persona). It MUST: (a) carry the style " & _
             "guidance and reviewer voice; (b) instruct the assistant to surface each " & _
             "recommendation WITH its strongest counter-case, anchored to an AR_ bookmark " & _
             "id, and to self-critique before finishing; (c) output a human-readable " & _
             "DECISION PACKET, never JSON. Do NOT include any JSONL output contract -- a " & _
             "separate cold serializer assistant owns that. Return ONLY the markdown code " & _
             "block containing the SKILL.md."

    modSysUtils.CopyToClipboard prompt

    MsgBox "Reduce Pass 3 Prompt copied to clipboard." & vbCrLf & vbCrLf & _
           "1. Paste this prompt into the same DHSChat conversation and send." & vbCrLf & _
           "2. This generates the HOT co-thinker SKILL.md. Save it via Save SKILL.md." & vbCrLf & _
           "3. The COLD serializer is set up once from TEMPLATE_SKILL_SERIALIZER.md " & _
           "(dashboard: Set Serializer URL).", vbInformation
End Sub

' Save SKILL.md
Public Sub SaveSkillMd()
    Dim activePersona As String
    Dim skillPath As String
    Dim fso As Object
    Dim ts As Object
    Dim skillContent As String
    Dim dataObj As Object
    
    activePersona = modAppCore.GetConfigValue("ActivePersona")
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
    
    modAppCore.UpsertPersona activePersona, skillMdPath:=skillPath
    
    MsgBox "Co-thinker SKILL.md saved to: " & skillPath & vbCrLf & vbCrLf & _
           "Next Step: Create a new DHSChat Assistant (the HOT co-thinker), paste this " & _
           "SKILL.md into its system prompt, and save its URL in the Personas sheet " & _
           "(AssistantUrl column). The shared COLD serializer is set up separately, once.", vbInformation
End Sub

