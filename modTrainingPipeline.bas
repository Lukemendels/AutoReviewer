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


' Read a revision's Author/Date without letting a dead COM object (e.g. a
' revision nested inside another deletion, or inside a field, that Word has
' already discarded by the time we get to it) abort the whole corpus run.
' Returns False -- and leaves outAuthor/outDate untouched -- if the object is
' unreadable; callers count that as one skipped record.
Private Function TryReadRevisionAuthorDate(ByVal rev As Object, ByRef outAuthor As String, ByRef outDate As Date) As Boolean
    On Error Resume Next
    Err.Clear
    outAuthor = rev.Author
    outDate = rev.Date
    TryReadRevisionAuthorDate = (Err.Number = 0)
    Err.Clear
End Function

' Read everything AddDocToCorpus needs from a target-author revision
' (Author/Date/Type/Range.Text) in one guarded pass. Returns False if any of
' them raise -- most commonly "Object has been deleted" on a revision whose
' range was invalidated by an earlier Accept or by an enclosing revision.
Private Function TryReadRevisionFull(ByVal rev As Object, ByRef outAuthor As String, ByRef outDate As Date, ByRef outType As Long, ByRef outText As String) As Boolean
    On Error Resume Next
    Err.Clear
    outAuthor = rev.Author
    outDate = rev.Date
    outType = rev.Type
    outText = rev.Range.Text
    TryReadRevisionFull = (Err.Number = 0)
    Err.Clear
End Function

' Comment counterparts of the above.
Private Function TryReadCommentAuthorDate(ByVal cmt As Object, ByRef outAuthor As String, ByRef outDate As Date) As Boolean
    On Error Resume Next
    Err.Clear
    outAuthor = cmt.Author
    outDate = cmt.Date
    TryReadCommentAuthorDate = (Err.Number = 0)
    Err.Clear
End Function

Private Function TryReadCommentFull(ByVal cmt As Object, ByRef outAuthor As String, ByRef outDate As Date, ByRef outText As String) As Boolean
    On Error Resume Next
    Err.Clear
    outAuthor = cmt.Author
    outDate = cmt.Date
    outText = cmt.Range.Text
    TryReadCommentFull = (Err.Number = 0)
    Err.Clear
End Function

Public Sub AddDocToCorpus()
    Const wdRevisionInsert As Long = 1
    Const wdRevisionDelete As Long = 2
    Const CONTEXT_CAP As Long = 500
    Const SCOPE_CAP As Long = 300

    Dim wdApp As Object
    Dim wdDoc As Object
    Dim docPath As String
    Dim targetAuthor As String
    Dim fd As FileDialog
    Dim screenUpdatingChanged As Boolean
    Dim paginationChanged As Boolean
    Dim origPagination As Boolean
    Dim completed As Boolean
    Dim recordsWritten As Long
    Dim skippedRecords As Long
    Dim scanAuthor As String
    Dim scanDate As Date

    Dim activePersona As String
    Dim personaToken As String

    On Error GoTo ErrHandler

    activePersona = Trim$(modAppCore.GetConfigValue("ActivePersona"))

    If activePersona = "" Then
        MsgBox "Please select an Active Persona first.", vbExclamation
        Exit Sub
    End If

    personaToken = modSysUtils.SafeFileToken(activePersona)
    If personaToken = "" Then Exit Sub

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
    On Error GoTo ErrHandler

    wdApp.Visible = True
    Set wdDoc = wdApp.Documents.Open(docPath)

    ' Word repaints on every revision Accept and Bookmark Add; with hundreds of
    ' revisions/paragraphs that repaint dominates run time. Suspend it for the
    ' whole pass and restore on every exit path (normal, early-return, error).
    wdApp.ScreenUpdating = False
    screenUpdatingChanged = True

    targetAuthor = modTrainingPipeline.SelectTargetAuthor(wdDoc)

    If targetAuthor = "" Then GoTo Cleanup

    ' Step: Compute target date range
    Dim minDate As Date
    Dim firstTargetFound As Boolean
    Dim rev As Object
    Dim cmt As Object

    firstTargetFound = False

    For Each rev In wdDoc.Revisions
        If TryReadRevisionAuthorDate(rev, scanAuthor, scanDate) Then
            If StrComp(scanAuthor, targetAuthor, vbTextCompare) = 0 Then
                If Not firstTargetFound Then
                    minDate = scanDate
                    firstTargetFound = True
                Else
                    If scanDate < minDate Then minDate = scanDate
                End If
            End If
        Else
            skippedRecords = skippedRecords + 1
        End If
    Next rev

    For Each cmt In wdDoc.Comments
        If TryReadCommentAuthorDate(cmt, scanAuthor, scanDate) Then
            If StrComp(scanAuthor, targetAuthor, vbTextCompare) = 0 Then
                If Not firstTargetFound Then
                    minDate = scanDate
                    firstTargetFound = True
                Else
                    If scanDate < minDate Then minDate = scanDate
                End If
            End If
        Else
            skippedRecords = skippedRecords + 1
        End If
    Next cmt

    ' Build baseline: Accept non-target revisions before target earliest.
    ' Accept removes the revision from wdDoc.Revisions and reindexes the
    ' collection, so we cannot enumerate-and-accept in one pass. Two-phase:
    ' For Each to collect the matching revisions into a VBA Collection (the
    ' Range/Revision references stay valid after other revisions are
    ' accepted), then Accept each one from that snapshot.
    Dim revToAccept As Collection
    Set revToAccept = New Collection
    For Each rev In wdDoc.Revisions
        If TryReadRevisionAuthorDate(rev, scanAuthor, scanDate) Then
            If StrComp(scanAuthor, targetAuthor, vbTextCompare) <> 0 Then
                If firstTargetFound And scanDate < minDate Then
                    revToAccept.Add rev
                End If
            End If
        Else
            skippedRecords = skippedRecords + 1
        End If
    Next rev

    ' Accepting hundreds of revisions one at a time triggers repagination and
    ' markup-display recalculation on each call; suspend both for the pass and
    ' restore unconditionally (Cleanup/ErrHandler).
    On Error Resume Next
    origPagination = wdApp.Options.Pagination
    wdApp.Options.Pagination = False
    paginationChanged = True
    wdDoc.Windows(1).View.ShowRevisionsAndComments = False
    On Error GoTo ErrHandler

    ' Accepting one revision can invalidate an adjacent paired revision (e.g.
    ' an insertion/deletion pair, or a revision nested inside one we just
    ' accepted) that is also sitting in this snapshot -- a failed Accept on an
    ' already-consumed revision is benign; count it and move on.
    Dim acceptRev As Object
    For Each acceptRev In revToAccept
        On Error Resume Next
        Err.Clear
        acceptRev.Accept
        If Err.Number <> 0 Then
            skippedRecords = skippedRecords + 1
            Err.Clear
        End If
        On Error GoTo ErrHandler
    Next acceptRev

    ' Stamp bookmarks. Extraction below relies on AR_PARA_ bookmarks for the
    ' per-record "anchor" field, so stamping MUST precede extraction.
    modWordUtils.StampDocWithArBookmarks wdDoc

    ' Extract records to JSONL
    Dim fso As Object
    Dim ts As Object
    Dim corpusPath As String

    corpusPath = modAppCore.GetPersonaFolder(activePersona) & "\" & personaToken & "_corpus.jsonl"

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(corpusPath, 8, True, -1) ' 8=Append, True=Create, -1=Unicode

    Dim docIdEsc As String
    Dim authorEsc As String
    docIdEsc = modSysUtils.JsonEscape(wdDoc.name)
    authorEsc = modSysUtils.JsonEscape(targetAuthor)

    ' Revisions output
    Dim jsonLine As String
    Dim revTypeStr As String
    Dim revAnchor As String
    Dim revContext As String
    Dim revParaIdx As Long
    Dim ctxRange As Object
    Dim revAuthorVal As String
    Dim revDateVal As Date
    Dim revTypeVal As Long
    Dim revTextVal As String

    For Each rev In wdDoc.Revisions
        If TryReadRevisionFull(rev, revAuthorVal, revDateVal, revTypeVal, revTextVal) Then
            If StrComp(revAuthorVal, targetAuthor, vbTextCompare) = 0 Then
                Select Case revTypeVal
                    Case wdRevisionInsert
                        revTypeStr = "insert"
                    Case wdRevisionDelete
                        revTypeStr = "delete"
                    Case Else
                        revTypeStr = "format/other"
                End Select

                ' AR_PARA numbering is a pure function of paragraph position (see
                ' modWordUtils.StampParagraphBookmarks), so the anchor for any
                ' range is computed directly -- one COM call, no bookmark scan.
                revAnchor = ""
                revContext = ""
                On Error Resume Next
                revParaIdx = wdDoc.Range(0, rev.Range.Start).Paragraphs.Count
                If revParaIdx > 0 Then
                    revAnchor = "AR_PARA_" & Format$(revParaIdx, "00000")
                End If
                Set ctxRange = rev.Range.Duplicate
                ctxRange.Expand 4 ' wdParagraph
                revContext = Trim$(ctxRange.Text)
                If Len(revContext) > CONTEXT_CAP Then revContext = Left$(revContext, CONTEXT_CAP)
                Set ctxRange = Nothing
                On Error GoTo ErrHandler

                jsonLine = "{""doc_id"":""" & docIdEsc & """" & _
                           ",""target_author"":""" & authorEsc & """" & _
                           ",""record_type"":""revision""" & _
                           ",""date"":""" & revDateVal & """" & _
                           ",""rev_type"":""" & revTypeStr & """" & _
                           ",""text"":""" & modSysUtils.JsonEscape(revTextVal) & """" & _
                           ",""anchor"":""" & revAnchor & """" & _
                           ",""context"":""" & modSysUtils.JsonEscape(revContext) & """}"
                ts.WriteLine jsonLine
                recordsWritten = recordsWritten + 1
            End If
        Else
            skippedRecords = skippedRecords + 1
        End If
    Next rev

    ' Comments output
    Dim cmtAnchor As String
    Dim cmtScopeText As String
    Dim cmtParaIdx As Long
    Dim cmtAuthorVal As String
    Dim cmtDateVal As Date
    Dim cmtTextVal As String

    For Each cmt In wdDoc.Comments
        If TryReadCommentFull(cmt, cmtAuthorVal, cmtDateVal, cmtTextVal) Then
            If StrComp(cmtAuthorVal, targetAuthor, vbTextCompare) = 0 Then
                cmtAnchor = ""
                cmtScopeText = ""
                On Error Resume Next
                cmtParaIdx = wdDoc.Range(0, cmt.Scope.Start).Paragraphs.Count
                If cmtParaIdx > 0 Then
                    cmtAnchor = "AR_PARA_" & Format$(cmtParaIdx, "00000")
                End If
                cmtScopeText = Trim$(cmt.Scope.Text)
                If Len(cmtScopeText) > SCOPE_CAP Then cmtScopeText = Left$(cmtScopeText, SCOPE_CAP)
                On Error GoTo ErrHandler

                jsonLine = "{""doc_id"":""" & docIdEsc & """" & _
                           ",""target_author"":""" & authorEsc & """" & _
                           ",""record_type"":""comment""" & _
                           ",""date"":""" & cmtDateVal & """" & _
                           ",""text"":""" & modSysUtils.JsonEscape(cmtTextVal) & """" & _
                           ",""scope_text"":""" & modSysUtils.JsonEscape(cmtScopeText) & """" & _
                           ",""anchor"":""" & cmtAnchor & """}"
                ts.WriteLine jsonLine
                recordsWritten = recordsWritten + 1
            End If
        Else
            skippedRecords = skippedRecords + 1
        End If
    Next cmt

    ts.Close

    ' Update Registry
    modAppCore.UpsertPersona activePersona, corpusPath:=corpusPath, incrementTrainingCount:=True

    ' Record the run -- including how many revisions/comments could not be
    ' read -- in the Trace sheet. Silent skips would hide a real gap in the
    ' persona's training data; a counted skip is honest lineage.
    On Error Resume Next
    modAudit.AppendReviewTrace "Corpus", activePersona, docPath, wdDoc.name, _
        "", "", "", recordsWritten + skippedRecords, recordsWritten, skippedRecords
    On Error GoTo ErrHandler

    completed = True

Cleanup:
    On Error Resume Next
    If paginationChanged And Not wdApp Is Nothing Then
        wdApp.Options.Pagination = origPagination
        paginationChanged = False
    End If
    If screenUpdatingChanged And Not wdApp Is Nothing Then
        wdApp.ScreenUpdating = True
        screenUpdatingChanged = False
    End If
    If Not wdDoc Is Nothing Then wdDoc.Close False
    Set wdDoc = Nothing
    Set wdApp = Nothing
    DoEvents
    On Error GoTo 0

    If completed Then
        MsgBox "Document added to corpus successfully." & vbCrLf & vbCrLf & _
               recordsWritten & " records written, " & skippedRecords & _
               " skipped (unreadable).", vbInformation
    End If
    Exit Sub

ErrHandler:
    MsgBox "Error adding document to corpus: " & Err.Description, vbCritical
    Resume Cleanup
End Sub


' Add a FINALIZED exemplar to the persona's training set.
' Instead of mining a messy redline diff, this captures the clean final text of
' a known-good document. Exemplars teach the *target state* (what good looks
' like); redlines teach the *transformation*. A persona can train on exemplars
' alone (the escape hatch when redlines are too messy -- multiple authors,
' overlapping turns) or on a mix of both. Saved as "<persona>_exemplar_NN.txt"
' beside the corpus and attached alongside it in the Reduce passes.
Public Sub AddFinalizedExemplar()
    Dim activePersona As String
    Dim personaToken As String
    Dim fd As FileDialog
    Dim docPath As String
    Dim wdApp As Object
    Dim wdDoc As Object
    Dim docName As String
    Dim cleanText As String
    Dim fso As Object
    Dim ts As Object
    Dim exPath As String
    Dim n As Long
    Dim screenUpdatingChanged As Boolean
    Dim workFolder As String

    On Error GoTo ErrHandler

    activePersona = Trim$(modAppCore.GetConfigValue("ActivePersona"))
    If activePersona = "" Then
        MsgBox "Please select an Active Persona first.", vbExclamation
        Exit Sub
    End If

    personaToken = modSysUtils.SafeFileToken(activePersona)
    If personaToken = "" Then Exit Sub

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "Select a FINALIZED Word document (a known-good example)"
    fd.Filters.Clear
    fd.Filters.Add "Word Documents", "*.docx"
    If fd.Show <> -1 Then Exit Sub
    docPath = fd.SelectedItems(1)

    On Error Resume Next
    Set wdApp = GetObject(, "Word.Application")
    If wdApp Is Nothing Then Set wdApp = CreateObject("Word.Application")
    On Error GoTo ErrHandler
    wdApp.Visible = True

    Set wdDoc = wdApp.Documents.Open(docPath)

    wdApp.ScreenUpdating = False
    screenUpdatingChanged = True

    ' Accept any stray revisions in memory so we capture the FINAL clean text.
    ' The file is closed without saving, so the source is never modified.
    On Error Resume Next
    wdDoc.TrackRevisions = False
    wdDoc.AcceptAllRevisions
    On Error GoTo ErrHandler

    docName = wdDoc.name
    cleanText = wdDoc.Content.Text

    wdApp.ScreenUpdating = True
    screenUpdatingChanged = False

    wdDoc.Close False
    Set wdDoc = Nothing
    Set wdApp = Nothing
    DoEvents

    ' Pick the next free exemplar filename for this persona.
    workFolder = modAppCore.GetPersonaFolder(activePersona)
    Set fso = CreateObject("Scripting.FileSystemObject")
    n = 1
    Do
        exPath = workFolder & "\" & personaToken & "_exemplar_" & Format(n, "00") & ".txt"
        If Not fso.FileExists(exPath) Then Exit Do
        n = n + 1
    Loop

    Set ts = fso.OpenTextFile(exPath, 2, True, -1) ' 2=Write, True=Create, -1=Unicode
    ts.Write "<<EXEMPLAR doc=""" & docName & """>>" & vbCrLf & cleanText
    ts.Close

    modAppCore.UpsertPersona activePersona, incrementExemplarCount:=True

    MsgBox "Finalized exemplar saved:" & vbCrLf & exPath & vbCrLf & vbCrLf & _
           "Add as many as you like, then run the Reduce passes. Exemplars are " & _
           "used as the gold standard for the persona's style -- with or without " & _
           "redlines from 'Add Doc to Corpus'.", vbInformation
    Exit Sub

ErrHandler:
    On Error Resume Next
    If screenUpdatingChanged And Not wdApp Is Nothing Then wdApp.ScreenUpdating = True
    If Not wdDoc Is Nothing Then wdDoc.Close False
    If Not wdApp Is Nothing Then Set wdApp = Nothing
    Set wdDoc = Nothing
    MsgBox "Error adding finalized exemplar: " & Err.Description, vbCritical
End Sub

' Pass 1: Cluster
Public Sub RunReducePass1()
    Dim activePersona As String
    Dim personaToken As String
    Dim workFolder As String
    Dim corpusPath As String
    Dim prompt As String
    Dim fso As Object
    Dim url As String
    Dim hasCorpus As Boolean
    Dim hasExemplars As Boolean
    Dim selectPath As String

    activePersona = Trim$(modAppCore.GetConfigValue("ActivePersona"))
    If activePersona = "" Then
        MsgBox "Please select an Active Persona first.", vbExclamation
        Exit Sub
    End If

    personaToken = modSysUtils.SafeFileToken(activePersona)
    If personaToken = "" Then Exit Sub

    workFolder = modAppCore.GetPersonaFolder(activePersona)
    corpusPath = workFolder & "\" & personaToken & "_corpus.jsonl"

    Set fso = CreateObject("Scripting.FileSystemObject")
    hasCorpus = fso.FileExists(corpusPath)
    hasExemplars = fso.FileExists(workFolder & "\" & personaToken & "_exemplar_01.txt")

    If Not hasCorpus And Not hasExemplars Then
        MsgBox "No training input found for this persona. Add redline docs " & _
               "(""Add Doc to Corpus"") and/or finalized exemplars " & _
               "(""Add Finalized Exemplar"") first.", vbExclamation
        Exit Sub
    End If

    ' One prompt covers redlines-only, exemplars-only, or both. Exemplars are the
    ' gold standard for what good looks like; redline records are observed edits.
    prompt = "I am attaching this reviewer's training material. It may include a " & _
             "corpus of their revisions/comments (observed edits) and/or one or more " & _
             "FINALIZED exemplar documents (their known-good output -- the gold " & _
             "standard for voice, structure, and standards). Infer the reviewer's " & _
             "implicit style preferences and cluster them into pattern categories. " & _
             "Where you have exemplars, learn what good looks like from them; where " & _
             "you have redlines, learn what they change and why."

    modSysUtils.CopyToClipboard prompt

    MsgBox "Reduce Pass 1 Prompt copied to clipboard." & vbCrLf & vbCrLf & _
           "1. Open a new chat in DHSChat." & vbCrLf & _
           "2. From the Explorer window that opens, drag in the corpus.jsonl AND " & _
           "any " & activePersona & "_exemplar_*.txt files for this persona." & vbCrLf & _
           "3. Paste the prompt and send.", vbInformation

    ' Open Explorer at a present file so the user can grab the corpus and exemplars.
    If hasCorpus Then
        selectPath = corpusPath
    Else
        selectPath = workFolder & "\" & personaToken & "_exemplar_01.txt"
    End If
    Shell "explorer.exe /select,""" & selectPath & """", vbNormalFocus

    ' Launch a fresh DHSChat chat (no assistant) in the browser. No hardcoded
    ' default: internal system URLs do not live in source.
    url = Trim$(modAppCore.GetConfigValue("CustomGptUrl", ""))
    If InStr(LCase(url), "http") > 0 Then
        Shell "explorer.exe """ & url & """", vbNormalFocus
    Else
        MsgBox "No CustomGptUrl is configured (Config sheet), so no browser was " & _
               "opened. Open DHSChat manually and start a new chat.", _
               vbExclamation, "No Chat URL"
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
             "guidance and reviewer voice; (b) define the three-turn protocol explicitly: " & _
             "Turn 1 THEMES -- cluster comments/revisions into 3-6 themes with a " & _
             "recommended posture and strongest counter-case per theme, then STOP and ask " & _
             "the human to rule on the themes before producing any blocks; Turn 2 BLOCKS " & _
             "-- only after the human rules, produce numbered '[n] BOOKMARK: <AR_ id> / " & _
             "ACTION / OLD_TEXT / NEW_TEXT / RATIONALE / COUNTER-CASE / CONFIDENCE' blocks " & _
             "consistent with the ratified themes, end with the COVERAGE line, then STOP " & _
             "and ask the human for a per-block KEEP / FIX: <instructions> / CUT ruling; " & _
             "Turn 3 FINAL RATIFIED PACKET -- only after the human rules on every block, " & _
             "reproduce KEEP blocks verbatim, apply FIX instructions exactly, omit CUT " & _
             "blocks WITHOUT renumbering the rest, drop COUNTER-CASE/CONFIDENCE from every " & _
             "surviving block, and make no other changes -- output only the final numbered " & _
             "blocks, nothing else; (c) instruct the assistant to self-critique before " & _
             "finishing each turn; (d) output human-readable DECISION PACKETS, never JSON, " & _
             "in all three turns. Do NOT include any JSONL output contract -- a separate " & _
             "cold serializer assistant owns that. Return ONLY the markdown code block " & _
             "containing the SKILL.md."

    modSysUtils.CopyToClipboard prompt

    MsgBox "Reduce Pass 3 Prompt copied to clipboard." & vbCrLf & vbCrLf & _
           "1. Paste this prompt into the same DHSChat conversation and send." & vbCrLf & _
           "2. This generates the HOT co-thinker SKILL.md (with the three-turn " & _
           "THEMES / BLOCKS / FINAL RATIFIED PACKET protocol). Save it via Save SKILL.md." & vbCrLf & _
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
    
    Dim personaToken As String

    activePersona = Trim$(modAppCore.GetConfigValue("ActivePersona"))
    If activePersona = "" Then
        MsgBox "Please select an Active Persona first.", vbExclamation
        Exit Sub
    End If

    personaToken = modSysUtils.SafeFileToken(activePersona)
    If personaToken = "" Then Exit Sub

    skillPath = modAppCore.GetPersonaFolder(activePersona) & "\" & personaToken & "_SKILL.md"
    
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

