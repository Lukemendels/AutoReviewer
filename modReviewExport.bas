Attribute VB_Name = "modReviewExport"
Option Explicit

Public Sub ExportWordDocForLLM(Optional ByVal isRespondMode As Boolean = False)
    Const wdMainTextStory As Long = 1
    Const wdSentence As Long = 3   ' Word's sentence unit constant
    
    Dim wdApp As Object          ' Word.Application
    Dim wdDoc As Object          ' Word.Document (single, sequentially modified in memory)
    Dim rngMain As Object        ' Word.Range
    Dim c As Object              ' Word.Comment
    Dim rScope As Object         ' Working range for scope sentence
    Dim fd As Object             ' FileDialog (late-bound)
    
    Dim wsConfig As Worksheet
    Dim wordPath As String
    Dim docsFolder As String
    Dim savePath As String
    Dim baseName As String
    Dim buffer As String
    Dim scopeSentence As String
    Dim highlightText As String
    Dim commentBody As String
    Dim stm As Object            ' ADODB.Stream for UTF-8
    Dim exportFormat As String

    On Error GoTo ErrHandler
    
    '---------------------------
    ' 1) Pick the Word document
    '---------------------------
    Set fd = Application.FileDialog(3) ' msoFileDialogFilePicker
    With fd
        .Title = "Select Word document to export for LLM"
        .AllowMultiSelect = False
        If Len(modAppCore.GetWorkFolder()) > 0 Then
            .InitialFileName = modAppCore.GetWorkFolder() & "\"
        End If
        
        If .Show <> -1 Then
            GoTo Cleanup   ' user cancelled
        End If
        
        wordPath = .SelectedItems(1)
    End With
    Set fd = Nothing
    
    wordPath = Trim$(wordPath)
    If Len(wordPath) = 0 Then GoTo Cleanup
    ' We trust Application.FileDialog to only return existing files.
    ' Do NOT call Dir$ on wordPath here; some path formats can cause
    ' "Bad file name or number" even when the file is valid.

    '---------------------------
    ' 1a) Reversibility hygiene: work on a copy, never the source of record.
    ' The user's selection is the source; we operate on "<base>_AR.<ext>" so the
    ' original is never stamped with bookmarks or written back into. Edits stay
    ' reversible (tracked changes on the copy) and the source stays pristine
    ' until the human chooses to finalize (Profile s7.2). If the copy cannot be
    ' made we fall back to operating on the original so the tool still runs.
    '---------------------------
    Dim sourcePath As String
    Dim workingPath As String
    sourcePath = wordPath
    workingPath = BuildWorkingCopyPath(sourcePath)

    On Error Resume Next
    Err.Clear
    If Len(workingPath) > 0 Then
        FileCopy sourcePath, workingPath
    End If
    If Err.Number <> 0 Or Len(workingPath) = 0 Then
        workingPath = sourcePath   ' fall back to the original
    End If
    On Error GoTo ErrHandler
    wordPath = workingPath

    '---------------------------
    ' 1b) Config sheet: record source and working paths via key-based helper.
    ' WordDocPath is the WORKING copy -- the apply step writes here.
    '---------------------------
    EnsureConfigSheet wsConfig
    SetConfigValue "SourceDocPath", sourcePath
    SetConfigValue "WordDocPath", workingPath
    
    ' Determine export format (plain | markdown)
    exportFormat = LCase$(Trim$(GetConfigValue("ExportFormat", "plain")))
    If exportFormat <> "markdown" Then
        exportFormat = "plain"
    End If
    
    '---------------------------
    ' 2) Start Word and open original doc (READ/WRITE)
    '---------------------------
    Set wdApp = CreateObject("Word.Application")
    ' Run Word hidden with screen updates off: this is a headless extraction that
    ' quits Word at the end regardless, and the per-paragraph/per-revision COM
    ' loops are far faster without repainting (the ~5-min export on heavily
    ' tracked docs was dominated by visible repaint).
    wdApp.Visible = False
    On Error Resume Next
    wdApp.ScreenUpdating = False
    wdApp.DisplayAlerts = 0
    On Error GoTo ErrHandler
    
    ' Open original document read/write so we can add bookmarks and SAVE them
    Set wdDoc = wdApp.Documents.Open(Filename:=wordPath, ReadOnly:=False)
    If wdDoc Is Nothing Then
        MsgBox "Word could not open the document: " & wordPath, vbCritical, "ExportWordDocForLLM"
        GoTo Cleanup
    End If
    
    '---------------------------
    ' 2a) STAMP DOC WITH AR BOOKMARKS (in memory only)
    '---------------------------
    ' Stamp the working copy so we can build the BOOKMARK_INDEX. We do NOT save
    ' these bookmarks: the apply step re-stamps the same (unchanged) working copy
    ' and gets identical ids, so the working copy left on disk stays a clean copy
    ' of the source. That way an abandoned export does not leave AR_ bookmarks
    ' lying around in the *_AR file.
    StampDocWithArBookmarks wdDoc
    If isRespondMode Then StampRevisionBookmarks wdDoc

    '---------------------------
    ' 3) Extract metadata (comments and revisions) BEFORE accepting revisions
    '---------------------------
    buffer = ""
    
    Dim bufferComments As String
    Dim bufferRevisions As String
    
    bufferComments = "<<COMMENTS_START>>" & vbCrLf

    ' Persist the full ordered list of comment ids so the apply step can
    ' warn-gate on any comment that received no reply or comment (Goal 1).
    Dim commentIdList As String
    commentIdList = ""

    If wdDoc.Comments.Count = 0 Then
        bufferComments = bufferComments & "(No comments in document)" & vbCrLf
    Else
        For Each c In wdDoc.Comments
            bufferComments = bufferComments & "## AR_COMMENT_" & c.Index & vbCrLf
            If Len(commentIdList) > 0 Then commentIdList = commentIdList & ","
            commentIdList = commentIdList & "AR_COMMENT_" & c.Index
            
            ' Author / Date
            On Error Resume Next
            bufferComments = bufferComments & "Author: " & CStr(c.Author) & vbCrLf
            bufferComments = bufferComments & "Date: " & CStr(c.Date) & vbCrLf
            On Error GoTo ErrHandler
            
            ' Compute Scope Sentence for this comment
            scopeSentence = ""
            highlightText = ""
            commentBody = ""
            
            On Error Resume Next
            highlightText = CStr(c.Scope.Text)   ' exact highlighted fragment
            commentBody = CStr(c.Range.Text)     ' comment text
            
            Set rScope = c.Scope.Duplicate
            If Not rScope Is Nothing Then
                rScope.Expand wdSentence         ' expand to full sentence(s)
                scopeSentence = CStr(rScope.Text)
            End If
            
            ' Fallback: if expansion fails or is empty, use the highlight itself
            If Len(Trim$(scopeSentence)) = 0 Then
                scopeSentence = highlightText
            End If
            On Error GoTo ErrHandler
            
            ' Output
            bufferComments = bufferComments & "Scope Sentence: " & CleanOneLine(scopeSentence) & vbCrLf
            bufferComments = bufferComments & "Highlight: " & CleanOneLine(highlightText) & vbCrLf
            bufferComments = bufferComments & "Text: " & CleanOneLine(commentBody) & vbCrLf
            
            bufferComments = bufferComments & "---" & vbCrLf
        Next c
    End If
    
    bufferComments = bufferComments & "<<COMMENTS_END>>" & vbCrLf & vbCrLf

    ' Record the comment-id list for the apply-side coverage warn-gate.
    SetConfigValue "CommentIds", commentIdList

    If isRespondMode Then
        bufferRevisions = "<<REVISIONS_START>>" & vbCrLf
        If wdDoc.Revisions.Count = 0 Then
            bufferRevisions = bufferRevisions & "(No revisions in document)" & vbCrLf
        Else
            Dim revObj As Object
            Dim revTypeStr As String
            Dim revIndexCounter As Long
            revIndexCounter = 1
            For Each revObj In wdDoc.Revisions
                bufferRevisions = bufferRevisions & "## AR_REV_" & Format(revIndexCounter, "00000") & vbCrLf
                
                On Error Resume Next
                bufferRevisions = bufferRevisions & "Author: " & CStr(revObj.Author) & vbCrLf
                bufferRevisions = bufferRevisions & "Date: " & CStr(revObj.Date) & vbCrLf
                
                ' Word Revision Types: wdRevisionInsert = 1, wdRevisionDelete = 2
                If revObj.Type = 1 Then
                    revTypeStr = "Insertion"
                ElseIf revObj.Type = 2 Then
                    revTypeStr = "Deletion"
                Else
                    revTypeStr = "Other (Type " & revObj.Type & ")"
                End If
                bufferRevisions = bufferRevisions & "Type: " & revTypeStr & vbCrLf
                
                ' Get the text (for deletion, it's the deleted text; for insertion, it's the new text)
                bufferRevisions = bufferRevisions & "Text: " & CleanOneLine(revObj.Range.Text) & vbCrLf
                On Error GoTo ErrHandler
                
                bufferRevisions = bufferRevisions & "---" & vbCrLf
                revIndexCounter = revIndexCounter + 1
            Next revObj
        End If
        bufferRevisions = bufferRevisions & "<<REVISIONS_END>>" & vbCrLf & vbCrLf
    End If

    '---------------------------
    ' 4) Accept All Revisions in memory to get "final" text, maintaining the original bookmark IDs
    '---------------------------
    On Error Resume Next
    wdDoc.TrackRevisions = False
    wdDoc.AcceptAllRevisions
    On Error GoTo ErrHandler
    
    '---------------------------
    ' 5) Build the export buffer (text, bookmarks, footnotes) from the finalized document
    '---------------------------
    ' DOCUMENT_TEXT (plain or markdown)
    buffer = buffer & BuildDocumentTextSection(Nothing, wdDoc, exportFormat)
    
    ' BOOKMARK_INDEX (paragraphs, table cells, footnotes)
    buffer = buffer & BuildBookmarkIndexSection(Nothing, wdDoc)
    
    ' FOOTNOTES section
    buffer = buffer & BuildFootnotesSection(Nothing, wdDoc, exportFormat)
    
    ' Append the pre-extracted comments and revisions
    buffer = buffer & bufferComments
    buffer = buffer & bufferRevisions

    '---------------------------
    ' 5a) Transport attestation (Profile s4.1 / s9.1): fingerprint the payload
    ' that will cross the clipboard, embed it self-describingly at the head of
    ' the artifact, and record it so the apply step can close the loop.
    '---------------------------
    Dim exportFingerprint As String
    exportFingerprint = modSysUtils.ArContentFingerprint(buffer)
    buffer = "<<PAYLOAD_FINGERPRINT: " & exportFingerprint & ">>" & vbCrLf & vbCrLf & buffer
    SetConfigValue "LastExportFingerprint", exportFingerprint
    SetConfigValue "LastExportMode", IIf(isRespondMode, "Respond", "Review")

    '---------------------------
    ' 6) Decide where to save .txt (UTF-8) - User's Documents
    '---------------------------
    docsFolder = GetUserDocumentsFolder()
    
    ' Fallbacks if Documents is not available
    If Len(docsFolder) = 0 Then
        docsFolder = modAppCore.GetWorkFolder()
        If Len(docsFolder) = 0 Then
            docsFolder = Environ$("USERPROFILE") & "\Documents"
        End If
    End If
    
    ' Trim and verify
    docsFolder = Trim$(docsFolder)
    If Len(docsFolder) = 0 Then
        Err.Raise vbObjectError + 100, "ExportWordDocForLLM", _
                  "Could not determine a valid folder for export."
    End If
    
    ' Ensure folder exists
    If Dir$(docsFolder, vbDirectory) = "" Then
        Err.Raise vbObjectError + 101, "ExportWordDocForLLM", _
                  "Export folder does not exist: " & docsFolder
    End If
    
    baseName = GetFileBaseName(wordPath)
    savePath = docsFolder & "\" & baseName & ".txt"

    '---------------------------
    ' 6a) Record LastExportTxtPath in Config via key-based helper
    '---------------------------
    SetConfigValue "LastExportTxtPath", savePath
    
    '---------------------------
    ' 7) Write buffer to file as UTF-8
    '---------------------------
    Set stm = CreateObject("ADODB.Stream")
    With stm
        .Type = 2            ' adTypeText
        .Charset = "UTF-8"   ' UTF-8 with BOM
        .Open
        .WriteText buffer
        .Position = 0
        .SaveToFile savePath, 2  ' adSaveCreateOverWrite
        .Close
    End With
    Set stm = Nothing
    
    '---------------------------
    ' 8) Sequential Teardown
    ' Close Word safely BEFORE showing Excel MsgBoxes or opening browsers
    '---------------------------
    On Error Resume Next
    If Not wdDoc Is Nothing Then wdDoc.Close SaveChanges:=False
    If Not wdApp Is Nothing Then
        wdApp.NormalTemplate.Saved = True
        wdApp.Quit SaveChanges:=False
    End If
    Set wdDoc = Nothing
    Set wdApp = Nothing
    DoEvents
    On Error GoTo ErrHandler
    
    ' 9) Setup Prompt and Browser
    ' HOT co-thinker prompt (Profile s7.4). This assistant runs divergent: it
    ' surfaces a *recommended* resolution with its *strongest counter-case* and
    ' a self-critique, citing the AR_ bookmark ID for each, and outputs a
    ' human-readable DECISION PACKET -- never JSONL. The human ratifies on paper
    ' (keep/fix/cut); the cold Serializer assistant converts the kept decisions
    ' to JSONL afterward. Keeping review and serialization in separate context
    ' windows is the s7.4 temperature wall.
    Dim promptText As String
    Dim packetSpec As String
    packetSpec = vbCrLf & vbCrLf & _
        "For each recommended change, output a numbered block in exactly this form:" & vbCrLf & _
        "[n] BOOKMARK: <the exact AR_ id, e.g. AR_PARA_00012 or AR_COMMENT_3>" & vbCrLf & _
        "    ACTION: replace_text | delete_element | add_comment_only | reply_to_comment | accept_revision | reject_revision" & vbCrLf & _
        "    OLD_TEXT: <the exact existing substring to change, or omit for a whole-bookmark/element action>" & vbCrLf & _
        "    NEW_TEXT: <the proposed replacement, when applicable>" & vbCrLf & _
        "    RATIONALE: <why>" & vbCrLf & _
        "    COUNTER-CASE: <the strongest argument AGAINST making this change>" & vbCrLf & _
        "    CONFIDENCE: High | Medium | Low" & vbCrLf & vbCrLf & _
        "Do NOT output JSON. Surface judgment for a human to ratify; a separate " & _
        "serializer will convert the ratified decisions to the edit format. Every " & _
        "block MUST cite a bookmark id that appears in the attached BOOKMARK_INDEX " & _
        "-- never invent an anchor. The bookmark id belongs ONLY on the BOOKMARK " & _
        "line: never write an AR_ id inside NEW_TEXT or a comment/reply body (they " & _
        "are internal anchors, not document content). Before you finish, " & _
        "self-critique: re-read each block and flag any that are unsupported by the " & _
        "source text."

    ' Respond mode leads with a middle-seat synthesis so the human grasps the
    ' feedback as a whole before adjudicating item by item.
    Dim synthesisSpec As String
    synthesisSpec = "First, before any per-item blocks, give a SYNTHESIS BRIEF:" & vbCrLf & _
        "  DIRECTION: in 2-4 sentences, where do these edits and comments push the document?" & vbCrLf & _
        "  THEMES: cluster the comments and revisions into a few named themes." & vbCrLf & _
        "  OPEN QUESTIONS: what must be answered to revise the document well?" & vbCrLf & _
        "Then the per-item blocks below." & vbCrLf

    ' Coverage requirement: every comment must be visibly adjudicated (Goal 1).
    Dim coverageSpec As String
    coverageSpec = vbCrLf & vbCrLf & _
        "COVERAGE REQUIREMENT: every AR_COMMENT_ id in the COMMENTS section MUST " & _
        "receive either a numbered block or an explicit NO_ACTION ruling with a " & _
        "one-line rationale. A comment left unmentioned is a failure. End your " & _
        "output with the line: COVERAGE: addressed <X> of <Y> comments; NO_ACTION: " & _
        "<ids or none>."

    If isRespondMode Then
        promptText = "I am attaching an exported Word document containing text, bookmark IDs, reviewer comments, and reviewer tracked changes (revisions). The text already reflects any accepted tracked changes. Follow your three-turn protocol: begin with Turn 1 (THEMES) and stop for my rulings before producing blocks. " & synthesisSpec & _
            "For each comment and each revision, unpack WHAT the reviewer is asking for, then surface the realistic OPTIONS -- accept as-is, modify, reject with rationale, or reply to the comment -- and recommend one with its counter-case. Anchor each to the reviewer's AR_COMMENT_ / AR_REV_ / paragraph id. EVERY reviewer comment MUST get a reply_to_comment block." & packetSpec & coverageSpec
    Else
        promptText = "I am attaching an exported Word document containing text and bookmark IDs. Follow your three-turn protocol: begin with Turn 1 (THEMES) and stop for my rulings before producing blocks." & packetSpec & coverageSpec
    End If
    modSysUtils.CopyToClipboard promptText
    
    ' Route selection. Synthetic review goes to the active persona's HOT
    ' co-thinker (style-specific). Incorporating supervisor feedback (Respond
    ' mode) goes to the single shared Incorporator -- understanding someone
    ' else's edits is style-agnostic, so it is infrastructure, not a persona
    ' (asymmetric model; mirrors the shared Serializer). No hardcoded default:
    ' internal system URLs do not live in source; configure via the dashboard
    ' or a CustomGptUrl config key.
    Dim gptUrl As String
    gptUrl = ""
    On Error Resume Next
    If isRespondMode Then
        gptUrl = Trim$(modAppCore.GetConfigValue("IncorporatorUrl", ""))
    Else
        Dim activePersona As String
        activePersona = modAppCore.GetConfigValue("ActivePersona")
        If Len(activePersona) > 0 Then gptUrl = Trim$(modAppCore.GetAssistantUrl(activePersona))
    End If
    If Len(gptUrl) = 0 Then gptUrl = Trim$(modAppCore.GetConfigValue("CustomGptUrl", ""))
    On Error GoTo 0
    ' Record the recommended route so the apply step's logic_trace can log
    ' recommended-vs-actual (Profile s9.2).
    SetConfigValue "LastRecommendedRoute", gptUrl
    If Len(gptUrl) > 0 Then
        modSysUtils.OpenURL gptUrl
    Else
        MsgBox "No assistant URL is configured, so no browser was opened." & vbCrLf & vbCrLf & _
               IIf(isRespondMode, _
                   "Set the shared Incorporator URL (dashboard: Set Incorporator URL).", _
                   "Set the active persona's AssistantUrl in the Personas sheet (or a CustomGptUrl key in Config).") & vbCrLf & vbCrLf & _
               "The export itself completed; open your assistant manually and continue.", _
               vbExclamation, "No Assistant URL"
    End If

    ' Export-time logic_trace row: abandoned reviews must still leave lineage,
    ' so the Trace sheet reflects the full run lifecycle, not only completed
    ' applies (apply columns stay blank until the apply run writes its row).
    On Error Resume Next
    modAudit.AppendReviewTrace _
        "Export", _
        GetConfigValue("ActivePersona", ""), _
        sourcePath, _
        wordPath, _
        gptUrl, _
        exportFingerprint, _
        "", 0, 0, 0
    On Error GoTo 0

    ' 10) Final Excel MsgBox (Word is already closed)
    Dim assistantLabel As String
    assistantLabel = IIf(isRespondMode, "Incorporator", "HOT co-thinker")
    MsgBox "Export complete (UTF-8):" & vbCrLf & savePath & vbCrLf & vbCrLf & _
           "The " & assistantLabel & " prompt is on your clipboard and the " & _
           assistantLabel & " assistant has been opened." & vbCrLf & vbCrLf & _
           "1. Paste the prompt (Ctrl+V) into the assistant." & vbCrLf & _
           "2. Upload/Drop the exported .txt file." & vbCrLf & _
           "3. Work through Turn 1 (THEMES) and Turn 2 (BLOCKS), ruling KEEP / " & _
           "FIX / CUT on each block." & vbCrLf & _
           "4. Paste the assistant's Turn 3 FINAL RATIFIED PACKET into the " & _
           "Ratified sheet at A8." & vbCrLf & _
           "5. Click ""Hand off to Serializer"", paste the prompt into the " & _
           "serializer chat, then paste the JSONL it returns into LLM_Changes!A8, " & _
           "then ""Apply LLM Edits to Word"".", vbInformation, "Export Complete"

    Exit Sub

Cleanup:
    On Error Resume Next
    If Not wdDoc Is Nothing Then wdDoc.Close SaveChanges:=False
    If Not wdApp Is Nothing Then
        wdApp.NormalTemplate.Saved = True
        wdApp.Quit SaveChanges:=False
    End If
    Set wdDoc = Nothing
    Set wdApp = Nothing
    Set fd = Nothing
    Set rScope = Nothing
    Set wsConfig = Nothing
    Exit Sub

ErrHandler:
    MsgBox "Error during export: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Public Sub ExportWordDocForRespondMode()
    ExportWordDocForLLM True
End Sub

' COLD serializer hand-off (Profile s7.4). After the human ratifies the hot
' co-thinker's decision packet, this copies the convergent serialize prompt and
' opens the shared Serializer assistant in a SEPARATE context window. The
' serializer's only job is serialize_exactly: translate the ratified decisions
' into the JSONL edit contract, never re-decide them.
'
' Completeness gate (item 8): the ratified packet pasted into the Ratified
' sheet (A8:A...) is the source of truth for WHICH anchors must come back from
' the serializer. Each surviving "[n] BOOKMARK: AR_..." line names one decision
' the serializer must translate; their count and ids are persisted to Config
' as ExpectedEditCount/ExpectedAnchors so ApplyWordSuggestionsFromJson can
' reconcile the pasted JSONL against them before any Word write.
Public Sub HandOffToSerializer()
    Dim serializerUrl As String
    Dim promptText As String
    Dim sessionToken As String
    Dim wb As Workbook
    Dim wsRatified As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim rawLines() As String
    Dim rawCount As Long
    Dim ratifiedText As String
    Dim anchors() As String
    Dim anchorCount As Long
    Dim expectedAnchorsCsv As String
    Dim i As Long

    Set wb = ThisWorkbook
    serializerUrl = Trim$(GetConfigValue("SerializerUrl", ""))

    ' Session binding: the serializer must echo the export fingerprint back in
    ' its meta line; the apply step verifies it before opening Word, so stale
    ' JSONL from another document cannot be applied. No export = no hand-off.
    sessionToken = Trim$(GetConfigValue("LastExportFingerprint", ""))
    If Len(sessionToken) = 0 Then
        MsgBox "No export fingerprint is recorded, so there is no session token " & _
               "to bind the serializer's output to." & vbCrLf & vbCrLf & _
               "Run the export step first, then hand off to the serializer.", _
               vbExclamation, "Hand Off to Serializer"
        Exit Sub
    End If

    ' Read the ratified packet (Turn 3 FINAL RATIFIED PACKET) from the Ratified
    ' sheet, A8:A..., one row per line.
    On Error Resume Next
    Set wsRatified = wb.Worksheets("Ratified")
    On Error GoTo 0
    If wsRatified Is Nothing Then
        modAppCore.SetupConfigAndLLMSheets
        On Error Resume Next
        Set wsRatified = wb.Worksheets("Ratified")
        On Error GoTo 0
    End If
    If wsRatified Is Nothing Then
        MsgBox "The Ratified sheet could not be found or created.", vbCritical, "Hand Off to Serializer"
        Exit Sub
    End If

    lastRow = wsRatified.Cells(wsRatified.Rows.Count, "A").End(xlUp).row
    If lastRow < 8 Then
        MsgBox "No ratified decisions found on the Ratified sheet (column A, starting at A8)." & vbCrLf & vbCrLf & _
               "Paste the HOT co-thinker's Turn 3 FINAL RATIFIED PACKET there first.", _
               vbExclamation, "Hand Off to Serializer"
        Exit Sub
    End If

    ReDim rawLines(1 To lastRow - 7)
    rawCount = 0
    For r = 8 To lastRow
        rawCount = rawCount + 1
        rawLines(rawCount) = CStr(wsRatified.Cells(r, "A").value)
    Next r

    ratifiedText = ""
    For i = 1 To rawCount
        If i > 1 Then ratifiedText = ratifiedText & vbCrLf
        ratifiedText = ratifiedText & rawLines(i)
    Next i

    If Len(Trim$(ratifiedText)) = 0 Then
        MsgBox "The Ratified sheet is empty." & vbCrLf & vbCrLf & _
               "Paste the HOT co-thinker's Turn 3 FINAL RATIFIED PACKET there first.", _
               vbExclamation, "Hand Off to Serializer"
        Exit Sub
    End If

    ' Each surviving decision keeps its "[n] BOOKMARK: AR_..." line (CUT
    ' decisions are omitted entirely, per the Turn 3 protocol). The count and
    ' ids of these lines are exactly what the serializer must return.
    anchorCount = ParseRatifiedAnchors(rawLines, rawCount, anchors)
    If anchorCount = 0 Then
        MsgBox "No ""[n] BOOKMARK: AR_..."" lines were found on the Ratified sheet." & vbCrLf & vbCrLf & _
               "Every kept or fixed decision must keep its bookmark line so the apply " & _
               "step can verify the serializer's output. Nothing was handed off.", _
               vbExclamation, "Hand Off to Serializer"
        Exit Sub
    End If

    expectedAnchorsCsv = ""
    For i = 1 To anchorCount
        If i > 1 Then expectedAnchorsCsv = expectedAnchorsCsv & ","
        expectedAnchorsCsv = expectedAnchorsCsv & anchors(i)
    Next i

    SetConfigValue "ExpectedEditCount", CStr(anchorCount)
    SetConfigValue "ExpectedAnchors", expectedAnchorsCsv

    promptText = "SESSION TOKEN: " & sessionToken & vbCrLf & vbCrLf & _
        "Below are ratified document-review decisions, each citing a " & _
        "bookmark id (AR_PARA_..., AR_CELL_..., AR_FN_..., or AR_COMMENT_...). " & _
        "Convert ONLY these decisions into the strict JSONL edit contract you " & _
        "hold, one JSON object per line. Translate exactly: do not re-decide, " & _
        "soften, reorder, add, or drop any decision, and do not change any " & _
        "bookmark id. Carry OLD_TEXT through verbatim as old_text when present " & _
        "so the edit stays surgical." & vbCrLf & vbCrLf & _
        "The ratified packet below contains exactly " & anchorCount & " decision(s). " & _
        "Your output MUST contain exactly " & anchorCount & " edit line(s) after the " & _
        "meta line -- one per decision, citing these bookmark ids: " & _
        expectedAnchorsCsv & ". If you cannot produce exactly that many edit lines, " & _
        "output ONLY an error line identifying the decision(s) you could not convert " & _
        "and why -- never emit a partial set." & vbCrLf & vbCrLf & _
        "The FIRST line of your output MUST be exactly this meta line, with N " & _
        "replaced by the number of edit lines that follow it:" & vbCrLf & _
        "{""meta"": ""autoreviewer"", ""session"": """ & sessionToken & """, ""count"": N}" & vbCrLf & vbCrLf & _
        "Output only the meta line and the JSONL lines -- no code fences, no " & _
        "commentary, nothing else." & _
        vbCrLf & vbCrLf & "RATIFIED DECISIONS:" & vbCrLf & ratifiedText

    modSysUtils.CopyToClipboard promptText

    If Len(serializerUrl) > 0 Then
        modSysUtils.OpenURL serializerUrl
        MsgBox "The COLD serializer prompt (with your " & anchorCount & " ratified " & _
               "decision(s) embedded) is on your clipboard and the Serializer " & _
               "assistant has been opened." & vbCrLf & vbCrLf & _
               "1. Paste the prompt." & vbCrLf & _
               "2. Copy the JSONL it returns into the LLM_Changes sheet at A8." & vbCrLf & _
               "3. Run ""Apply LLM Edits to Word""." & vbCrLf & vbCrLf & _
               "The apply step will verify the JSONL covers exactly these " & _
               anchorCount & " anchor(s): " & expectedAnchorsCsv, vbInformation, "Hand Off to Serializer"
    Else
        MsgBox "The COLD serializer prompt (with your " & anchorCount & " ratified " & _
               "decision(s) embedded) is on your clipboard." & vbCrLf & vbCrLf & _
               "No Serializer URL is set yet. Set one via the dashboard " & _
               "(""Set Serializer URL"") so this step can open it automatically. " & _
               "For now, open your Serializer assistant manually and paste the prompt.", _
               vbExclamation, "Hand Off to Serializer"
    End If
End Sub

'=== Helpers ========================================================

' Scan the ratified packet for surviving "[n] BOOKMARK: <id>" lines (the Turn 3
' protocol keeps these verbatim for KEEP/FIX decisions and omits CUT decisions
' entirely). Returns the count and fills anchors(1..count) with the bookmark
' ids in the order they appear.
Private Function ParseRatifiedAnchors(ByRef lines() As String, ByVal n As Long, _
                                      ByRef anchors() As String) As Long
    Dim i As Long
    Dim m As Long
    Dim t As String
    Dim bracketPos As Long
    Dim rest As String
    Dim anchorId As String

    ReDim anchors(1 To IIf(n > 0, n, 1))
    m = 0
    For i = 1 To n
        t = modReviewImport.TrimWs(lines(i))
        If Left$(t, 1) = "[" Then
            bracketPos = InStr(t, "]")
            If bracketPos > 1 Then
                rest = modReviewImport.TrimWs(Mid$(t, bracketPos + 1))
                If UCase$(Left$(rest, 9)) = "BOOKMARK:" Then
                    anchorId = modReviewImport.TrimWs(Mid$(rest, 10))
                    If Len(anchorId) > 0 Then
                        m = m + 1
                        anchors(m) = anchorId
                    End If
                End If
            End If
        End If
    Next i
    ParseRatifiedAnchors = m
End Function

Private Function BuildBookmarkIndexSection(ByVal wdDocFinal As Object, _
                                           ByVal wdDoc As Object) As String
    Dim docForExport As Object
    Dim bm As Object
    Dim sb As String
    Dim name As String
    Dim t As String
    Dim snippet As String
    
    On Error GoTo SafeExit
    
    ' Prefer FINAL doc (revisions accepted) for bookmark ranges
    If Not wdDocFinal Is Nothing Then
        Set docForExport = wdDocFinal
    Else
        Set docForExport = wdDoc
    End If
    
    If docForExport Is Nothing Then GoTo SafeExit
    
    ' If there are no bookmarks, emit an empty section
    If docForExport.Bookmarks.Count = 0 Then
        sb = "<<BOOKMARK_INDEX_START>>" & vbCrLf & _
             "(No AR bookmarks found in document)" & vbCrLf & _
             "<<BOOKMARK_INDEX_END>>" & vbCrLf & vbCrLf
        BuildBookmarkIndexSection = sb
        Exit Function
    End If
    
    sb = "<<BOOKMARK_INDEX_START>>" & vbCrLf
    
    For Each bm In docForExport.Bookmarks
        On Error Resume Next
        name = CStr(bm.name)
        On Error GoTo SafeExit
        
        ' Only include our AR_* bookmarks. Prefix lengths are exact:
        ' "AR_" = 3, "AR_PARA_"/"AR_CELL_" = 8, "AR_FN_" = 6.
        If Left$(name, 3) = "AR_" Then
            ' Classify type based on prefix
            If Left$(name, 8) = "AR_PARA_" Then
                t = "paragraph"
            ElseIf Left$(name, 8) = "AR_CELL_" Then
                t = "table_cell"
            ElseIf Left$(name, 6) = "AR_FN_" Then
                t = "footnote"
            Else
                t = "other"
            End If
            
            On Error Resume Next
            snippet = CleanOneLine(bm.Range.Text)
            On Error GoTo SafeExit
            
            If Len(snippet) > 200 Then
                snippet = Left$(snippet, 200) & "..."
            End If
            
            sb = sb & name & " | type=" & t & " | """ & snippet & """" & vbCrLf
        End If
    Next bm
    
    sb = sb & "<<BOOKMARK_INDEX_END>>" & vbCrLf & vbCrLf
    
    BuildBookmarkIndexSection = sb
    Exit Function
    
SafeExit:
    BuildBookmarkIndexSection = ""
End Function
Private Function BuildFootnotesSection(ByVal wdDocFinal As Object, _
                                       ByVal wdDoc As Object, _
                                       ByVal exportFormat As String) As String
    Dim docForExport As Object
    Dim sectionText As String
    Dim fn As Object
    Dim anchorSnippet As String
    Dim fnBody As String
    Dim paraText As String
    Dim fmt As String
    
    On Error GoTo SafeExit
    
    ' Use the FINAL doc (revisions accepted) if available, otherwise the original
    If Not wdDocFinal Is Nothing Then
        Set docForExport = wdDocFinal
    Else
        Set docForExport = wdDoc
    End If
    
    If docForExport Is Nothing Then GoTo SafeExit
    
    ' If no footnotes, nothing to add
    On Error Resume Next
    If docForExport.Footnotes.Count = 0 Then GoTo SafeExit
    On Error GoTo SafeExit
    
    exportFormat = LCase$(exportFormat)
    sectionText = "<<FOOTNOTES_START>>" & vbCrLf
    
    For Each fn In docForExport.Footnotes
        ' Anchor snippet: paragraph containing the footnote reference
        On Error Resume Next
        paraText = ""
        If Not fn.Reference Is Nothing Then
            paraText = fn.Reference.Paragraph.Range.Text
        End If
        On Error GoTo SafeExit
        
        anchorSnippet = CleanOneLine(paraText)
        If Len(anchorSnippet) > 200 Then
            anchorSnippet = Left$(anchorSnippet, 200) & "..."
        End If
        
        ' Footnote body text
        On Error Resume Next
        fnBody = CleanOneLine(fn.Range.Text)
        On Error GoTo SafeExit
        
        If Len(fnBody) > 0 Or Len(anchorSnippet) > 0 Then
            If exportFormat = "markdown" Then
                sectionText = sectionText & "## Footnote " & fn.Index & vbCrLf
                sectionText = sectionText & "Anchor: " & anchorSnippet & vbCrLf
                sectionText = sectionText & "Text: " & fnBody & vbCrLf
                sectionText = sectionText & "---" & vbCrLf
            Else
                sectionText = sectionText & "Footnote " & fn.Index & vbCrLf
                sectionText = sectionText & "Anchor: " & anchorSnippet & vbCrLf
                sectionText = sectionText & "Text: " & fnBody & vbCrLf
                sectionText = sectionText & "----" & vbCrLf
            End If
        End If
    Next fn
    
    sectionText = sectionText & "<<FOOTNOTES_END>>" & vbCrLf & vbCrLf
    BuildFootnotesSection = sectionText
    Exit Function
    
SafeExit:
    BuildFootnotesSection = ""
End Function


Private Function BuildDocumentTextSection(ByVal wdDocFinal As Object, _
                                          ByVal wdDoc As Object, _
                                          ByVal exportFormat As String) As String
    Dim docForExport As Object
    Dim docText As String
    Dim para As Object
    Dim rawText As String
    Dim cleaned As String
    Dim styleName As String
    Dim headingLevel As Long
    Dim isList As Boolean
    Dim line As String
    Dim lastWasBlank As Boolean
    
    On Error GoTo FallbackPlain
    
    '---------------------------
    ' Decide which document to use for export
    ' Prefer the FINAL doc (revisions accepted)
    '---------------------------
    If Not wdDocFinal Is Nothing Then
        Set docForExport = wdDocFinal
    Else
        Set docForExport = wdDoc
    End If
    
    If docForExport Is Nothing Then GoTo FallbackPlain
    
    ' Normalize exportFormat
    exportFormat = LCase$(exportFormat)
    If exportFormat <> "markdown" Then
        exportFormat = "plain"
    End If
    
    '---------------------------
    ' Plain export (no Markdown)
    '---------------------------
    If exportFormat = "plain" Then
        On Error Resume Next
        docText = modSysUtils.NormalizePunctuation(docForExport.Content.Text)
        On Error GoTo FallbackPlain
        
        BuildDocumentTextSection = "<<DOCUMENT_TEXT_START>>" & vbCrLf & _
                                   docText & _
                                   "<<DOCUMENT_TEXT_END>>" & vbCrLf & vbCrLf
        Exit Function
    End If
    
    '---------------------------
    ' Markdown export
    '---------------------------
    docText = ""
    lastWasBlank = False
    
    ' Diagnostic: see what Word thinks is in the doc
    On Error Resume Next
    Debug.Print "BuildDocumentTextSection:", _
                "exportFormat=" & exportFormat & "; " & _
                "ContentLen=" & Len(docForExport.Content.Text) & "; " & _
                "Paras=" & docForExport.Paragraphs.Count
    On Error GoTo FallbackPlain
    
    For Each para In docForExport.Paragraphs
        On Error Resume Next
        rawText = para.Range.Text
        On Error GoTo FallbackPlain
        
        cleaned = CleanOneLine(rawText)
        
        ' Skip purely empty paragraphs, but preserve spacing
        If Len(cleaned) = 0 Then
            If Not lastWasBlank And Len(docText) > 0 Then
                docText = docText & vbCrLf
                lastWasBlank = True
            End If
            GoTo NextParagraph
        End If
        
        ' Determine style name (e.g., "Heading 1", "Heading 2", etc.)
        styleName = ""
        On Error Resume Next
        styleName = CStr(para.Style)
        On Error GoTo FallbackPlain
        
        headingLevel = 0
        Select Case LCase$(styleName)
            Case "heading 1"
                headingLevel = 1
            Case "heading 2"
                headingLevel = 2
            Case "heading 3"
                headingLevel = 3
            Case "heading 4"
                headingLevel = 4
        End Select
        
        ' Determine if this paragraph is part of a list
        isList = False
        On Error Resume Next
        If Not para.Range Is Nothing Then
            If Not para.Range.ListFormat Is Nothing Then
                If para.Range.ListFormat.ListType <> 0 Then
                    isList = True
                End If
            End If
        End If
        On Error GoTo FallbackPlain
        
        ' Build Markdown line
        If headingLevel > 0 Then
            ' Heading: #, ##, ###, etc.
            line = String$(headingLevel, "#") & " " & cleaned
        ElseIf isList Then
            ' Treat all lists as bullet lists for now: "- item"
            line = "- " & cleaned
        Else
            ' Regular paragraph
            line = cleaned
        End If
        
        ' Add blank line between blocks if needed
        If Len(docText) > 0 And Not lastWasBlank Then
            docText = docText & vbCrLf
        End If
        
        docText = docText & line & vbCrLf
        lastWasBlank = False
        
NextParagraph:
    Next para
    
    BuildDocumentTextSection = "<<DOCUMENT_TEXT_START>>" & vbCrLf & _
                               docText & _
                               "<<DOCUMENT_TEXT_END>>" & vbCrLf & vbCrLf
    Exit Function
    
FallbackPlain:
    ' If anything goes wrong, fall back to simple plain export using Content.Text
    On Error Resume Next
    If wdDocFinal Is Nothing Then
        Set docForExport = wdDoc
    Else
        Set docForExport = wdDocFinal
    End If
    
    If Not docForExport Is Nothing Then
        docText = modSysUtils.NormalizePunctuation(docForExport.Content.Text)
    Else
        docText = ""
    End If
    On Error GoTo 0
    
    BuildDocumentTextSection = "<<DOCUMENT_TEXT_START>>" & vbCrLf & _
                               docText & _
                               "<<DOCUMENT_TEXT_END>>" & vbCrLf & vbCrLf
End Function

' Build the working-copy path "<dir>\<base>_AR.<ext>" alongside the source.
' Returns "" if the path can't be parsed, signalling the caller to fall back
' to operating on the original.
Private Function BuildWorkingCopyPath(ByVal sourcePath As String) As String
    Dim slashPos As Long
    Dim dotPos As Long
    Dim dir As String
    Dim fileName As String
    Dim baseName As String
    Dim ext As String

    slashPos = InStrRev(sourcePath, "\")
    If slashPos = 0 Then
        BuildWorkingCopyPath = ""
        Exit Function
    End If

    dir = Left$(sourcePath, slashPos)            ' includes trailing backslash
    fileName = Mid$(sourcePath, slashPos + 1)

    dotPos = InStrRev(fileName, ".")
    If dotPos > 0 Then
        baseName = Left$(fileName, dotPos - 1)
        ext = Mid$(fileName, dotPos)             ' includes the dot
    Else
        baseName = fileName
        ext = ""
    End If

    ' Avoid stacking "_AR_AR" if the user re-selects a working copy.
    If Right$(baseName, 3) = "_AR" Then
        BuildWorkingCopyPath = dir & baseName & ext
    Else
        BuildWorkingCopyPath = dir & baseName & "_AR" & ext
    End If
End Function

Private Function GetFileBaseName(ByVal fullPath As String) As String
    Dim filePart As String
    Dim dotPos As Long
    Dim invalidChars As Variant
    Dim i As Long
    Dim ch As String
    
    ' Extract filename without path
    filePart = Mid$(fullPath, InStrRev(fullPath, "\") + 1)
    
    ' Strip extension
    dotPos = InStrRev(filePart, ".")
    If dotPos > 0 Then
        filePart = Left$(filePart, dotPos - 1)
    End If
    
    ' Replace characters that are invalid in Windows filenames
    invalidChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For i = LBound(invalidChars) To UBound(invalidChars)
        filePart = Replace(filePart, invalidChars(i), "_")
    Next i
    
    ' Trim trailing dots or spaces (also invalid)
    Do While Len(filePart) > 0 And (Right$(filePart, 1) = " " Or Right$(filePart, 1) = ".")
        filePart = Left$(filePart, Len(filePart) - 1)
    Loop
    
    ' Fallback if everything got stripped
    If Len(filePart) = 0 Then
        filePart = "Export"
    End If
    
    GetFileBaseName = filePart
End Function


' Returns the user's Documents folder (late-bound, no references).
Private Function GetUserDocumentsFolder() As String
    Dim wsh As Object
    On Error Resume Next
    Set wsh = CreateObject("WScript.Shell")
    If Not wsh Is Nothing Then
        GetUserDocumentsFolder = CStr(wsh.SpecialFolders("MyDocuments"))
    Else
        GetUserDocumentsFolder = ""
    End If
    Set wsh = Nothing
End Function

Private Function CleanOneLine(ByVal s As String) As String
    Dim tmp As String
    tmp = s

    ' Normalize smart/Unicode punctuation to ASCII so the payload can't render as
    ' box-question-mark "tofu" downstream (and round-trips cleanly on apply).
    tmp = modSysUtils.NormalizePunctuation(tmp)

    ' Replace line breaks with spaces
    tmp = Replace(tmp, vbCr, " ")
    tmp = Replace(tmp, vbLf, " ")
    
    ' Collapse multiple spaces
    Do While InStr(tmp, "  ") > 0
        tmp = Replace(tmp, "  ", " ")
    Loop
    
    CleanOneLine = Trim$(tmp)
End Function

Private Function EnsureTxtExtension(ByVal p As String) As String
    Dim dotPos As Long
    Dim slashPos As Long
    Dim base As String
    
    dotPos = InStrRev(p, ".")
    slashPos = InStrRev(p, "\")
    
    If dotPos > slashPos Then
        base = Left$(p, dotPos - 1)
    Else
        base = p
    End If
    
    EnsureTxtExtension = base & ".txt"
End Function

