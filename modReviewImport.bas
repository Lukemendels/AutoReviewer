Attribute VB_Name = "modReviewImport"
Option Explicit

' Bookmark-only, JSONL-driven Word editing
' - Anchors exclusively by bookmark_id (e.g., AR_PARA_00037, AR_CELL_1_2_3, AR_FN_001)
' - No para_index or comment_index; no old_text locators

Public Sub ApplyWordSuggestionsFromJson()
    Const CHANGE_REPLACE_TEXT   As String = "replace_text"
    Const CHANGE_DELETE_ELEMENT As String = "delete_element"
    Const CHANGE_ADD_COMMENT    As String = "add_comment_only"
    Const CHANGE_REPLY_COMMENT  As String = "reply_to_comment"
    Const CHANGE_ACCEPT_REVISION As String = "accept_revision"
    Const CHANGE_REJECT_REVISION As String = "reject_revision"
    Const CHANGE_ADD_FOOTNOTE    As String = "add_footnote"

    ' Late-bound equivalents for Final / No Markup view (kept for consistency)
    Const wdRevisionsMarkupNone As Long = 0   ' hide all markup
    Const wdRevisionsViewFinal  As Long = 0   ' show final view
    
    Dim wb As Workbook
    Dim wsConfig As Worksheet
    Dim wsChanges As Worksheet
    Dim wsLog As Worksheet
    Dim logRow As Long
    
    Dim wordPath As String
    Dim lines() As String
    Dim line As String
    Dim i As Long
    
    Dim wdApp As Object  ' Word.Application
    Dim wdDoc As Object  ' Word.Document
    Dim targetRange As Object

    ' Tracked-change provenance: edits are stamped as "AutoReviewer" so AI
    ' suggestions are visibly attributable and one-click-rejectable. UserName is
    ' an application-global Word setting, so we snapshot and MUST restore it on
    ' every exit path or the operator's real Word name stays overwritten.
    Dim origUserName As String
    Dim origInitials As String
    Dim userNameChanged As Boolean
    
    Dim bookmarkId As String
    Dim changeType As String
    Dim oldText As String
    Dim newText As String
    Dim addComment As String
    Dim applyChange As Variant
    Dim applyThis As Boolean
    Dim confidenceRaw As String
    Dim confidenceNorm As String
    
    Dim totalLines As Long
    Dim parsedOk As Long
    Dim appliedCount As Long
    Dim skippedCount As Long

    ' Provenance/completeness tracking
    Dim madeTextEdit As Boolean         ' did we create at least one tracked text edit
    Dim origCommentCount As Long        ' reviewer comments present at open
    Dim repliedComments As Object       ' Scripting.Dictionary of replied comment indices
    Dim revAuthorSeen As String         ' read-back: author Word actually stamped

    ' Config-driven behavior
    Dim defaultConfidenceLevel As String
    Dim useArPrefix As Boolean
    
    ' For reading JSONL from sheet
    Dim lastRow As Long
    Dim r As Long
    Dim tmp As String
    Dim tmpLines() As String
    Dim n As Long
    
    On Error GoTo ErrHandler
    
    Set wb = ThisWorkbook
    
    '---------------------------
    ' 1) Get WordDocPath & AR config from Config (key-based)
    '---------------------------
    EnsureConfigSheet wsConfig
    wordPath = Trim$(GetConfigValue("WordDocPath", ""))
    
    If Len(wordPath) = 0 Then
        MsgBox "WordDocPath (Config key 'WordDocPath') is empty. " & _
               "Run the export macro to set it, or fill it manually in the Config sheet.", _
               vbExclamation
        GoTo Cleanup
    End If
    
    ' For local paths only, do a Dir$ check; URLs are allowed without Dir$
    If InStr(1, wordPath, "://", vbTextCompare) = 0 Then
        If Dir$(wordPath, vbNormal) = "" Then
            MsgBox "The Word document path from the Config sheet does not exist:" & vbCrLf & _
                   wordPath, vbCritical
            GoTo Cleanup
        End If
    End If
    
    ' Read AR-related config flags
    defaultConfidenceLevel = NormalizeConfidence(GetConfigValue("DefaultConfidenceLevel", "Medium"), "Medium")
    useArPrefix = GetConfigBool("UseArCommentPrefix", False)
    
    '---------------------------
    ' 2) Get JSONL from LLM_Changes (A8:A...)
    '---------------------------
    On Error Resume Next
    Set wsChanges = wb.Worksheets("LLM_Changes")
    On Error GoTo ErrHandler
    
    If wsChanges Is Nothing Then
        MsgBox "LLM_Changes sheet not found. Run SetupLLMWorkflowSheets first.", vbExclamation
        GoTo Cleanup
    End If
    
    lastRow = wsChanges.Cells(wsChanges.Rows.Count, "A").End(xlUp).row
    If lastRow < 8 Then
        MsgBox "No JSONL lines found in column A starting at A8.", vbExclamation
        GoTo Cleanup
    End If
    
    Dim rawLines() As String
    Dim rawCount As Long
    ReDim rawLines(1 To lastRow - 7)
    rawCount = 0
    For r = 8 To lastRow
        rawCount = rawCount + 1
        rawLines(rawCount) = CStr(wsChanges.Cells(r, "A").value)
    Next r

    ' Fence-tolerant payload extraction: if the serializer's fenced ```jsonl
    ' block is present, take only the lines inside it (prose notes after the
    ' closing fence are ignored); otherwise take all non-blank lines. The
    ' operator may paste with or without fences -- both gate identically.
    n = FilterPayloadLines(rawLines, rawCount, tmpLines)

    If n = 0 Then
        MsgBox "No JSONL payload found in column A starting at A8 " & _
               "(a fence-only paste counts as empty).", vbExclamation
        GoTo Cleanup
    End If

    '---------------------------
    ' 2a) Session-binding gate (default-deny, before Word opens).
    ' Bookmark ids are generic ordinals, so a stale payload from a previous
    ' document can apply cleanly to the WRONG document. The serializer's first
    ' line must be a meta line carrying the export fingerprint and the edit
    ' count; any mismatch aborts with no partial apply.
    '---------------------------
    Dim sessionFailCode As String
    If Not CheckSessionGate(tmpLines, n, Trim$(GetConfigValue("LastExportFingerprint", "")), sessionFailCode) Then
        MsgBox "Session gate blocked the apply (" & sessionFailCode & "):" & vbCrLf & vbCrLf & _
               SessionFailMessage(sessionFailCode) & vbCrLf & vbCrLf & _
               "Nothing was applied.", vbCritical, "Session Binding"
        GoTo Cleanup
    End If

    ' Drop the meta line; lines() carries only the edit lines.
    If n = 1 Then
        MsgBox "Session gate passed, but the payload contains zero edit lines.", _
               vbInformation, "Nothing To Apply"
        GoTo Cleanup
    End If
    ReDim lines(1 To n - 1)
    For r = 2 To n
        lines(r - 1) = tmpLines(r)
    Next r

    ' Transport attestation (Profile s9.1): fingerprint the exact JSONL block
    ' the operator pasted (meta line included), so the logic_trace records the
    ' bytes that produced the edits.
    Dim jsonlFingerprint As String
    Dim joinedJsonl As String
    Dim k As Long
    For k = 1 To n
        joinedJsonl = joinedJsonl & tmpLines(k) & vbLf
    Next k
    jsonlFingerprint = modSysUtils.ArContentFingerprint(joinedJsonl)

    '---------------------------
    ' 2c) Parse and validate every edit line ONCE, up front and with no Word
    ' open. The two-pass apply below reuses these arrays; the coverage gate
    ' needs them too, and parsing before Word opens means an Abort never spins
    ' up Word.
    '---------------------------
    totalLines = UBound(lines) - LBound(lines) + 1
    Application.StatusBar = "Parsing suggestions: " & totalLines & " lines"

    Dim pParsed() As Boolean
    Dim pValCode() As String
    Dim pIsRev() As Boolean
    Dim pBookmark() As String
    Dim pChange() As String
    Dim pOld() As String
    Dim pNew() As String
    Dim pComment() As String
    Dim pApply() As Variant
    Dim pConf() As String
    ReDim pParsed(1 To totalLines)
    ReDim pValCode(1 To totalLines)
    ReDim pIsRev(1 To totalLines)
    ReDim pBookmark(1 To totalLines)
    ReDim pChange(1 To totalLines)
    ReDim pOld(1 To totalLines)
    ReDim pNew(1 To totalLines)
    ReDim pComment(1 To totalLines)
    ReDim pApply(1 To totalLines)
    ReDim pConf(1 To totalLines)

    For i = 1 To totalLines
        pParsed(i) = ParseJsonLine(lines(i), bookmarkId, changeType, oldText, newText, _
                                   addComment, applyChange, confidenceRaw)
        If pParsed(i) Then
            pBookmark(i) = bookmarkId
            pChange(i) = changeType
            pOld(i) = oldText
            pNew(i) = newText
            pComment(i) = addComment
            pApply(i) = applyChange
            pConf(i) = confidenceRaw
            pValCode(i) = ValidateParsedChange(bookmarkId, changeType, oldText, newText, addComment)
            If pValCode(i) = "" Then
                Dim ctL As String
                ctL = LCase$(TrimWs(changeType))
                pIsRev(i) = (ctL = CHANGE_ACCEPT_REVISION Or ctL = CHANGE_REJECT_REVISION)
            End If
        End If
    Next i

    '---------------------------
    ' 2d) Comment-coverage warn-gate (Profile s9.4 false-negative guard). The
    ' export persisted the full ordered list of AR_COMMENT_ ids; a comment that
    ' received neither a reply_to_comment nor an add_comment_only is unaddressed.
    ' This is a WARN, not a hard block: a ratifier may rule no-action, but that
    ' must be a visible choice. The decision is logged to the Trace row.
    '---------------------------
    Dim coverageDecision As String
    Dim unaddressedList As String
    Dim covIds() As String
    Dim covCount As Long
    coverageDecision = "n/a"
    unaddressedList = ""
    covCount = SplitCsv(GetConfigValue("CommentIds", ""), covIds)
    If covCount > 0 Then
        unaddressedList = ComputeUnaddressed(covIds, covCount, pChange, pBookmark, totalLines)
        If Len(unaddressedList) = 0 Then
            coverageDecision = "All addressed"
        Else
            Dim covResp As VbMsgBoxResult
            covResp = MsgBox("These reviewer comments received no reply and no comment:" & vbCrLf & vbCrLf & _
                             Replace(unaddressedList, ",", vbCrLf) & vbCrLf & vbCrLf & _
                             "Proceed anyway (recording a no-action ruling), or " & _
                             "Abort to revise the decisions first?", _
                             vbExclamation + vbYesNo + vbDefaultButton2, "Unaddressed Comments")
            If covResp = vbNo Then
                coverageDecision = "Abort"
                On Error Resume Next
                modAudit.AppendReviewTrace "Apply-Aborted", GetConfigValue("ActivePersona", ""), _
                    GetConfigValue("SourceDocPath", ""), wordPath, _
                    GetConfigValue("LastRecommendedRoute", ""), _
                    GetConfigValue("LastExportFingerprint", ""), jsonlFingerprint, _
                    totalLines, 0, 0, unaddressedList, coverageDecision
                On Error GoTo ErrHandler
                GoTo Cleanup
            Else
                coverageDecision = "Proceed despite unaddressed"
            End If
        End If
    End If
    
    '---------------------------
    ' 2a) Ensure Log sheet exists (bookmark_id-focused)
    '---------------------------
    On Error Resume Next
    Set wsLog = wb.Worksheets("Log")
    On Error GoTo ErrHandler
    
    If wsLog Is Nothing Then
        Set wsLog = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        wsLog.name = "Log"
        With wsLog
            .Range("A1").value = "Timestamp"
            .Range("B1").value = "LineNumber"
            .Range("C1").value = "BookmarkId"
            .Range("D1").value = "ChangeType"
            .Range("E1").value = "Status"
            .Range("F1").value = "Reason"
            .Range("G1").value = "JsonLine"
            .Range("H1").value = "Confidence"
            .Range("I1").value = "Pass"
            .Columns("A:I").EntireColumn.AutoFit
        End With
    Else
        ' Ensure headers are aligned with bookmark_id schema
        If CStr(wsLog.Range("C1").value) <> "BookmarkId" Then wsLog.Range("C1").value = "BookmarkId"
        If CStr(wsLog.Range("H1").value) <> "Confidence" Then wsLog.Range("H1").value = "Confidence"
        If CStr(wsLog.Range("I1").value) <> "Pass" Then wsLog.Range("I1").value = "Pass"
        wsLog.Columns("A:I").EntireColumn.AutoFit
    End If
    
    logRow = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).row + 1
    
    '---------------------------
    ' 3) Open Word and document
    '---------------------------
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = True
    On Error Resume Next
    wdApp.DisplayAlerts = 0   ' wdAlertsNone

    ' Snapshot and override the revision author for the duration of the apply.
    origUserName = CStr(wdApp.UserName)
    origInitials = CStr(wdApp.UserInitials)
    wdApp.UserName = "AutoReviewer"
    wdApp.UserInitials = "AR"
    userNameChanged = True
    On Error GoTo ErrHandler

    Set wdDoc = Nothing
    On Error Resume Next
    Set wdDoc = wdApp.Documents.Open(Filename:=wordPath, ReadOnly:=False)
    On Error GoTo ErrHandler
    
    If wdDoc Is Nothing Then
        MsgBox "Word could not open the document at:" & vbCrLf & wordPath, vbCritical
        GoTo Cleanup
    End If
    
    ' Configure Final / No Markup view so we see final text
    On Error Resume Next
    With wdApp.ActiveWindow.View.RevisionsFilter
        .Markup = wdRevisionsMarkupNone
        .View = wdRevisionsViewFinal
    End With
    On Error GoTo ErrHandler

    ' Re-stamp the AR_ anchors. The export deliberately does not persist them, so
    ' an abandoned *_AR working copy stays clean; we reproduce them here. Because
    ' this is the identical (unchanged) working copy the export stamped, the ids
    ' match what the model saw in the BOOKMARK_INDEX. Done before TrackRevisions
    ' so the stamping itself is never recorded as a revision.
    StampDocWithArBookmarks wdDoc
    StampRevisionBookmarks wdDoc

    wdDoc.TrackRevisions = True

    ' Snapshot reviewer comments present at open, and prepare the replied-to set,
    ' for the every-comment-addressed completeness check.
    origCommentCount = wdDoc.Comments.Count
    Set repliedComments = CreateObject("Scripting.Dictionary")
    madeTextEdit = False

    '---------------------------
    ' 4) Apply in TWO passes over the pre-parsed set: text/comment edits first,
    ' accept/reject_revision second. Accepting or rejecting a revision can
    ' delete ranges that other bookmarks live inside, so revision verdicts must
    ' never run before the text edits they could invalidate. The Log sheet
    ' records the pass number per line.
    '---------------------------
    Dim passNum As Long
    For passNum = 1 To 2

    For i = 1 To totalLines
        ' Pass routing: revision verdicts are deferred to pass 2; everything
        ' else (including parse/validation failures) is handled in pass 1.
        If passNum = 1 And pIsRev(i) Then GoTo SkipLine
        If passNum = 2 And Not pIsRev(i) Then GoTo SkipLine

        Dim logStatus As String
        Dim logReason As String

        line = lines(i)
        logStatus = ""
        logReason = ""
        Dim commentTarget As Object
        applyThis = True
        Set targetRange = Nothing
        Set commentTarget = Nothing

        Application.StatusBar = "Applying suggestions (pass " & passNum & " of 2): line " & i & " of " & totalLines
        DoEvents

        If Not pParsed(i) Then
            logStatus = "Skipped"
            logReason = "ParseJsonLine failed"
            skippedCount = skippedCount + 1
            GoTo LogAndNext
        End If

        parsedOk = parsedOk + 1

        ' Load this line's parsed fields
        bookmarkId = pBookmark(i)
        changeType = pChange(i)
        oldText = pOld(i)
        newText = pNew(i)
        addComment = pComment(i)
        applyChange = pApply(i)
        confidenceRaw = pConf(i)
        confidenceNorm = NormalizeConfidence(confidenceRaw, defaultConfidenceLevel)

        ' Respect apply_change flag (boolean false skips the line)
        applyThis = True
        If Not IsEmpty(applyChange) Then
            If VarType(applyChange) = vbBoolean Then applyThis = applyChange
        End If

        If Not applyThis Then
            logStatus = "Skipped"
            logReason = "apply_change=false"
            skippedCount = skippedCount + 1
            GoTo LogAndNext
        End If

        ' Contract validation (shared with the self-test harness)
        If pValCode(i) <> "" Then
            logStatus = "Skipped"
            logReason = ValidationMessage(pValCode(i), changeType)
            skippedCount = skippedCount + 1
            GoTo LogAndNext
        End If

        ' Locate bookmark or comment
        ' "AR_COMMENT_" is 11 chars; the index begins at position 12.
        If Left$(bookmarkId, 11) = "AR_COMMENT_" Then
            Dim cIndex As Long
            cIndex = Val(Mid$(bookmarkId, 12))
            On Error Resume Next
            Set commentTarget = wdDoc.Comments(cIndex)
            On Error GoTo ErrHandler
            
            If commentTarget Is Nothing Then
                logStatus = "Skipped"
                logReason = "Comment not found: " & bookmarkId
                skippedCount = skippedCount + 1
                GoTo LogAndNext
            End If
        Else
            On Error Resume Next
            Set targetRange = wdDoc.Bookmarks(bookmarkId).Range
            On Error GoTo ErrHandler
            
            If targetRange Is Nothing Then
                logStatus = "Skipped"
                logReason = "Bookmark not found: " & bookmarkId
                skippedCount = skippedCount + 1
                GoTo LogAndNext
            End If
        End If
        
        '---------------------------
        ' Apply operation based on change_type
        '---------------------------
        Select Case LCase$(TrimWs(changeType))
            Case CHANGE_REPLACE_TEXT
                ' Anchor ids must never land in the document text.
                newText = modSysUtils.StripArTokens(newText)

                Dim editRange As Object
                Dim txtRange As String
                Dim subPos As Long
                Dim s0 As Long
                Dim hayNorm As String
                Dim ndlNorm As String
                Dim changed As Boolean

                ' Work on a copy of the bookmark range
                Set editRange = targetRange.Duplicate

                ' For paragraph bookmarks, exclude the trailing paragraph mark
                ' so we don't delete the separator between this paragraph and the next.
                ' "AR_PARA_" is 8 chars.
                If Left$(bookmarkId, 8) = "AR_PARA_" Then
                    txtRange = CStr(editRange.Text)
                    If Len(txtRange) > 0 And Right$(txtRange, 1) = Chr$(13) Then
                        editRange.End = editRange.End - 1   ' trim off the paragraph mark
                    End If
                End If

                ' If old_text is supplied, narrow the range to that span first so the
                ' diff is scoped to the intended occurrence. Match on the normalized
                ' text so a smart-quote/dash in the doc matches a straight one from the
                ' model. Normalization is 1:1, so positions map onto real offsets.
                If Len(oldText) > 0 Then
                    hayNorm = modSysUtils.NormalizePunctuation(CStr(editRange.Text))
                    ndlNorm = modSysUtils.NormalizePunctuation(oldText)
                    subPos = InStr(1, hayNorm, ndlNorm, vbBinaryCompare)
                    If subPos = 0 Then
                        logStatus = "Skipped"
                        logReason = "old_text not found within bookmark: " & bookmarkId
                        skippedCount = skippedCount + 1
                        GoTo LogAndNext
                    End If
                    s0 = editRange.Start
                    editRange.End = s0 + (subPos - 1) + Len(ndlNorm)
                    editRange.Start = s0 + (subPos - 1)
                End If

                ' Surgical write-back: track-change only the differing middle, so a
                ' one-word change is a one-word revision, not a whole-paragraph
                ' delete+insert.
                changed = ReplaceMinimalSpan(editRange, newText)

                If changed Then
                    appliedCount = appliedCount + 1
                    madeTextEdit = True
                    logStatus = "Applied"
                    logReason = "replace_text"
                    If Len(addComment) > 0 Then
                        AddArComment wdDoc, editRange, addComment, confidenceNorm, useArPrefix
                    End If
                Else
                    logStatus = "Skipped"
                    logReason = "replace_text: new_text equals existing text"
                    skippedCount = skippedCount + 1
                End If
           
            Case CHANGE_DELETE_ELEMENT
                ' Delete entire range
                targetRange.Text = ""
                madeTextEdit = True

                ' Re-add a zero-length bookmark at this location so the ID remains valid
                On Error Resume Next
                wdDoc.Bookmarks.Add name:=bookmarkId, Range:=targetRange
                On Error GoTo ErrHandler

                appliedCount = appliedCount + 1
                logStatus = "Applied"
                logReason = "delete_element"

                If Len(addComment) > 0 Then
                    AddArComment wdDoc, targetRange, addComment, confidenceNorm, useArPrefix
                End If

            Case CHANGE_ADD_COMMENT
                AddArComment wdDoc, targetRange, addComment, confidenceNorm, useArPrefix
                appliedCount = appliedCount + 1
                logStatus = "Applied"
                logReason = "add_comment_only"

            Case CHANGE_REPLY_COMMENT
                ' Validation guarantees an AR_COMMENT_ target with reply text;
                ' the locate step above guarantees commentTarget resolved.
                AddArCommentReply wdDoc, commentTarget, addComment, confidenceNorm, useArPrefix

                ' Record that this reviewer comment was addressed (completeness).
                On Error Resume Next
                repliedComments(CStr(cIndex)) = True
                On Error GoTo ErrHandler

                appliedCount = appliedCount + 1
                logStatus = "Applied"
                logReason = "reply_to_comment"

            Case CHANGE_ACCEPT_REVISION
                If targetRange.Revisions.Count > 0 Then
                    Dim rAccept As Object
                    For Each rAccept In targetRange.Revisions
                        rAccept.Accept
                    Next rAccept
                    appliedCount = appliedCount + 1
                    logStatus = "Applied"
                    logReason = "accept_revision"
                Else
                    logStatus = "Skipped"
                    logReason = "No revisions found in target range"
                    skippedCount = skippedCount + 1
                End If
                
            Case CHANGE_REJECT_REVISION
                If targetRange.Revisions.Count > 0 Then
                    Dim rReject As Object
                    For Each rReject In targetRange.Revisions
                        rReject.Reject
                    Next rReject
                    appliedCount = appliedCount + 1
                    logStatus = "Applied"
                    logReason = "reject_revision"
                Else
                    logStatus = "Skipped"
                    logReason = "No revisions found in target range"
                    skippedCount = skippedCount + 1
                End If

            Case CHANGE_ADD_FOOTNOTE
                ' Insert a footnote (the citation body in new_text) at the end of
                ' the target range, or immediately after old_text when given. The
                ' callout/footnote is a tracked insertion authored AutoReviewer.
                Dim fnBody As String
                Dim fnRange As Object
                Dim fnTxt As String
                Dim fnPos As Long

                fnBody = modSysUtils.StripArTokens(newText)
                Set fnRange = targetRange.Duplicate

                ' Trim trailing paragraph mark for paragraph anchors.
                If Left$(bookmarkId, 8) = "AR_PARA_" Then
                    fnTxt = CStr(fnRange.Text)
                    If Len(fnTxt) > 0 And Right$(fnTxt, 1) = Chr$(13) Then
                        fnRange.End = fnRange.End - 1
                    End If
                End If

                If Len(oldText) > 0 Then
                    ' Place the callout immediately after the matched span.
                    hayNorm = modSysUtils.NormalizePunctuation(CStr(fnRange.Text))
                    ndlNorm = modSysUtils.NormalizePunctuation(oldText)
                    fnPos = InStr(1, hayNorm, ndlNorm, vbBinaryCompare)
                    If fnPos = 0 Then
                        logStatus = "Skipped"
                        logReason = "add_footnote: old_text not found in " & bookmarkId
                        skippedCount = skippedCount + 1
                        GoTo LogAndNext
                    End If
                    fnRange.Start = fnRange.Start + (fnPos - 1) + Len(ndlNorm)
                End If

                ' Collapse to an insertion point at the end of the (located) range.
                fnRange.Start = fnRange.End

                On Error Resume Next
                wdDoc.Footnotes.Add Range:=fnRange, Text:=fnBody
                If Err.Number <> 0 Then
                    On Error GoTo ErrHandler
                    logStatus = "Skipped"
                    logReason = "add_footnote: insertion failed at " & bookmarkId
                    skippedCount = skippedCount + 1
                    GoTo LogAndNext
                End If
                On Error GoTo ErrHandler

                madeTextEdit = True
                appliedCount = appliedCount + 1
                logStatus = "Applied"
                logReason = "add_footnote"

            Case Else
                logStatus = "Skipped"
                logReason = "Unknown change_type: " & changeType
                skippedCount = skippedCount + 1
        End Select
        
LogAndNext:
        ' Default status if somehow not set
        If Len(logStatus) = 0 Then
            logStatus = "Skipped"
            If Len(logReason) = 0 Then logReason = "No operation performed"
        End If

        ' Write log entry (includes bookmark_id, Confidence, and pass number)
        If Not wsLog Is Nothing Then
            On Error Resume Next
            wsLog.Cells(logRow, 1).value = Now
            wsLog.Cells(logRow, 2).value = i   ' edit line number (after meta)
            wsLog.Cells(logRow, 3).value = bookmarkId
            wsLog.Cells(logRow, 4).value = changeType
            wsLog.Cells(logRow, 5).value = logStatus
            wsLog.Cells(logRow, 6).value = logReason
            wsLog.Cells(logRow, 7).value = line
            wsLog.Cells(logRow, 8).value = confidenceNorm
            wsLog.Cells(logRow, 9).value = passNum
            logRow = logRow + 1
            On Error GoTo ErrHandler
        End If

SkipLine:
    Next i

    Next passNum

    ' Terminal step: strip all AR_ anchors so the delivered document is clean and
    ' a future pass re-stamps from scratch (re-run hygiene). Runs after every
    ' edit/comment is applied, before the save.
    On Error Resume Next
    RemoveArBookmarks wdDoc
    On Error GoTo ErrHandler

    ' Read-back author diagnostic: did our text edits actually get authored
    ' "AutoReviewer"? On account-signed-in Word, revision author can follow the
    ' signed-in account regardless of Application.UserName -- this reports what
    ' actually stuck, so the operator isn't guessing.
    revAuthorSeen = ""
    If madeTextEdit Then
        On Error Resume Next
        Dim rvChk As Object
        For Each rvChk In wdDoc.Revisions
            If StrComp(CStr(rvChk.Author), "AutoReviewer", vbTextCompare) = 0 Then
                revAuthorSeen = "AutoReviewer"
                Exit For
            End If
        Next rvChk
        On Error GoTo ErrHandler
    End If

    Dim oldBgSave As Boolean

    '---------------------------
    ' 5) Save and summarize
    '---------------------------
    On Error Resume Next
    oldBgSave = wdApp.Options.BackgroundSave
    wdApp.Options.BackgroundSave = False
    wdDoc.Save
    wdApp.Options.BackgroundSave = oldBgSave
    On Error GoTo ErrHandler

    ' 5a) The run completed and the document is saved: clear the pasted payload
    ' so a stale JSONL cannot linger in LLM_Changes and be re-applied to the
    ' wrong document later (belt to the session gate's suspenders).
    On Error Resume Next
    wsChanges.Range("A8:A" & lastRow).ClearContents
    On Error GoTo ErrHandler

    '---------------------------
    ' 6) Sequential Teardown
    ' Close Word safely BEFORE showing Excel MsgBoxes
    '---------------------------
    On Error Resume Next
    Application.StatusBar = False
    ' Restore the operator's Word author name before quitting (it is global).
    If userNameChanged And Not wdApp Is Nothing Then
        wdApp.UserName = origUserName
        wdApp.UserInitials = origInitials
        userNameChanged = False
    End If
    If Not wdDoc Is Nothing Then wdDoc.Close SaveChanges:=True
    If Not wdApp Is Nothing Then
        wdApp.NormalTemplate.Saved = True
        wdApp.Quit SaveChanges:=False
    End If
    Set wdDoc = Nothing
    Set wdApp = Nothing
    DoEvents
    On Error GoTo ErrHandler
    
    ' 6a) Append the run's logic_trace (Profile s9.3). This is the defensible
    ' artifact: who ran it, the recommended route, and the transport
    ' fingerprints linking export payload -> pasted JSONL -> edits.
    On Error Resume Next
    modAudit.AppendReviewTrace _
        GetConfigValue("LastExportMode", "Review"), _
        GetConfigValue("ActivePersona", ""), _
        GetConfigValue("SourceDocPath", ""), _
        wordPath, _
        GetConfigValue("LastRecommendedRoute", ""), _
        GetConfigValue("LastExportFingerprint", ""), _
        jsonlFingerprint, _
        totalLines, appliedCount, skippedCount, _
        unaddressedList, coverageDecision
    On Error GoTo ErrHandler

    ' 7) Show Final Summary
    Dim summaryMsg As String
    Dim unaddressed As Long

    summaryMsg = "JSONL lines: " & totalLines & vbCrLf & _
                 "Parsed OK: " & parsedOk & vbCrLf & _
                 "Applied: " & appliedCount & vbCrLf & _
                 "Skipped: " & skippedCount & vbCrLf & vbCrLf & _
                 "JSONL fingerprint: " & jsonlFingerprint & vbCrLf & _
                 "Logged to the Trace sheet."

    ' Completeness: were all reviewer comments replied to?
    If origCommentCount > 0 Then
        unaddressed = origCommentCount - repliedComments.Count
        If unaddressed < 0 Then unaddressed = 0
        summaryMsg = summaryMsg & vbCrLf & vbCrLf & _
                     "Reviewer comments: " & origCommentCount & _
                     " | replied: " & repliedComments.Count & _
                     IIf(unaddressed > 0, " | UNADDRESSED: " & unaddressed, " | all addressed")
    End If

    ' Author provenance read-back.
    If madeTextEdit Then
        If revAuthorSeen = "AutoReviewer" Then
            summaryMsg = summaryMsg & vbCrLf & "Revision author: AutoReviewer (applied)."
        Else
            summaryMsg = summaryMsg & vbCrLf & _
                         "Revision author: NOT AutoReviewer -- Word used the signed-in " & _
                         "account. Comments are still authored AutoReviewer. See the guide " & _
                         "to force insertion author."
        End If
    End If

    MsgBox summaryMsg, vbInformation, "Apply Bookmark-Based Suggestions"
    
    Exit Sub

Cleanup:
    On Error Resume Next
    Application.StatusBar = False
    ' Restore the operator's Word author name on the error/early-exit path too.
    If userNameChanged And Not wdApp Is Nothing Then
        wdApp.UserName = origUserName
        wdApp.UserInitials = origInitials
        userNameChanged = False
    End If
    ' Cleanup is reached ONLY by an early GoTo (before any edits) or by
    ' ErrHandler (a fault mid-apply). The successful run saves and closes in its
    ' own teardown block and never lands here. So we close WITHOUT saving: a
    ' half-applied document must not be persisted -- that would contradict the
    ' session gate's no-partial-apply posture. The unsaved working copy is
    ' discarded; the operator re-runs.
    If Not wdDoc Is Nothing Then wdDoc.Close SaveChanges:=False
    If Not wdApp Is Nothing Then
        wdApp.NormalTemplate.Saved = True
        wdApp.Quit SaveChanges:=False
    End If
    Set wdDoc = Nothing
    Set wdApp = Nothing
    Set wsConfig = Nothing
    Set wsChanges = Nothing
    Set wsLog = Nothing
    Exit Sub

ErrHandler:
    MsgBox "Error applying suggestions: " & Err.Description & vbCrLf & vbCrLf & _
           "No changes were saved; the document is unchanged. Re-run after " & _
           "fixing the cause.", vbCritical, "Apply Aborted"
    Resume Cleanup
End Sub

'=== JSONL tokenizer (twin: ref/jsonl_contract.py) ===================
' A single left-to-right pass over the line that walks string literals
' (escape-aware) and recognizes keys only at object top level. Replaces the
' old InStr-based extractors and fixes two defects:
'   1. Escape parity: a closing quote is recognized iff it is preceded by an
'      EVEN run of backslashes (the reader consumes backslash+next as a unit;
'      the old prevCh check misread "a\\" as unterminated).
'   2. Key-in-value collision: a key name appearing inside another field's
'      VALUE can no longer be mistaken for the key.
' Semantics shared with the Python twin -- never change one side alone:
' TrimWs strips space/tab/CR/LF both ends; duplicate keys: FIRST wins; value
' types are tagged s=string (unescaped) / b=bool / n=number (raw) / z=null /
' c=complex (raw nested object or array).

' Trim that also strips CR/LF (VBA Trim$ strips spaces only). Public because
' the self-test harness shares it.
Public Function TrimWs(ByVal s As String) As String
    Dim a As Long, b As Long
    Dim ch As String
    a = 1
    b = Len(s)
    Do While a <= b
        ch = Mid$(s, a, 1)
        If ch = " " Or ch = vbTab Or ch = vbCr Or ch = vbLf Then a = a + 1 Else Exit Do
    Loop
    Do While b >= a
        ch = Mid$(s, b, 1)
        If ch = " " Or ch = vbTab Or ch = vbCr Or ch = vbLf Then b = b - 1 Else Exit Do
    Loop
    If b >= a Then TrimWs = Mid$(s, a, b - a + 1)
End Function

' A code-fence line: any trimmed line beginning with three backticks (so a
' language tag like ```jsonl or ```JSONL is recognized).
Public Function IsFenceLine(ByVal trimmed As String) As Boolean
    IsFenceLine = (Left$(trimmed, 3) = "```")
End Function

' Collect the payload lines from a paste (twin: filter_payload_lines in
' ref/session.py). The serializer emits its meta+edit lines inside a single
' ```jsonl fenced block, with refuse-don't-guess notes (prose) AFTER the
' closing fence. So: if any fence line is present, take only the lines strictly
' between the FIRST fence and the next fence, dropping blanks; everything
' outside the fenced block (including stray prose after it) is ignored. If no
' fence is present, the operator pasted raw JSONL: take all non-blank lines.
' Fills outLines(1..count) and returns count.
Public Function FilterPayloadLines(ByRef rawLines() As String, ByVal rawCount As Long, _
                                   ByRef outLines() As String) As Long
    Dim i As Long
    Dim t As String
    Dim startF As Long
    Dim endF As Long
    Dim lastIdx As Long
    Dim m As Long
    Dim trimmed() As String

    If rawCount <= 0 Then
        FilterPayloadLines = 0
        Exit Function
    End If

    ReDim trimmed(1 To rawCount)
    For i = 1 To rawCount
        trimmed(i) = TrimWs(rawLines(i))
    Next i

    startF = 0
    For i = 1 To rawCount
        If IsFenceLine(trimmed(i)) Then
            startF = i
            Exit For
        End If
    Next i

    ReDim outLines(1 To rawCount)
    m = 0

    If startF > 0 Then
        endF = 0
        For i = startF + 1 To rawCount
            If IsFenceLine(trimmed(i)) Then
                endF = i
                Exit For
            End If
        Next i
        If endF > 0 Then lastIdx = endF - 1 Else lastIdx = rawCount
        For i = startF + 1 To lastIdx
            t = trimmed(i)
            If Len(t) > 0 And Not IsFenceLine(t) Then
                m = m + 1
                outLines(m) = t
            End If
        Next i
    Else
        For i = 1 To rawCount
            t = trimmed(i)
            If Len(t) > 0 Then
                m = m + 1
                outLines(m) = t
            End If
        Next i
    End If

    FilterPayloadLines = m
End Function

' Split a comma-separated string into outItems(1..count); returns count.
' Empty / whitespace items are dropped. Used for the persisted CommentIds list.
Public Function SplitCsv(ByVal s As String, ByRef outItems() As String) As Long
    Dim parts() As String
    Dim i As Long
    Dim t As String
    Dim m As Long

    If Len(TrimWs(s)) = 0 Then
        SplitCsv = 0
        Exit Function
    End If

    parts = Split(s, ",")
    ReDim outItems(1 To UBound(parts) - LBound(parts) + 1)
    m = 0
    For i = LBound(parts) To UBound(parts)
        t = TrimWs(parts(i))
        If Len(t) > 0 Then
            m = m + 1
            outItems(m) = t
        End If
    Next i
    SplitCsv = m
End Function

Private Function SkipWs(ByVal s As String, ByVal i As Long) As Long
    Dim ch As String
    Do While i <= Len(s)
        ch = Mid$(s, i, 1)
        If ch = " " Or ch = vbTab Or ch = vbCr Or ch = vbLf Then i = i + 1 Else Exit Do
    Loop
    SkipWs = i
End Function

' s position i is the opening quote. Returns True with the RAW contents
' (escapes intact) and iNext just past the closing quote; False if
' unterminated. A backslash consumes the following character, so a quote
' closes the string iff preceded by an even backslash run.
Private Function ReadJsonStringRaw(ByVal s As String, ByVal i As Long, _
                                   ByRef raw As String, ByRef iNext As Long) As Boolean
    Dim n As Long
    Dim ch As String
    Dim sb As String
    n = Len(s)
    i = i + 1
    sb = ""
    Do While i <= n
        ch = Mid$(s, i, 1)
        If ch = "\" Then
            If i + 1 > n Then Exit Function   ' dangling backslash: unterminated
            sb = sb & ch & Mid$(s, i + 1, 1)
            i = i + 2
        ElseIf ch = """" Then
            raw = sb
            iNext = i + 1
            ReadJsonStringRaw = True
            Exit Function
        Else
            sb = sb & ch
            i = i + 1
        End If
    Loop
End Function

' s position i is "{" or "[": walk (string-aware) to the matching close.
Private Function ReadJsonNestedRaw(ByVal s As String, ByVal i As Long, _
                                   ByRef raw As String, ByRef iNext As Long) As Boolean
    Dim n As Long
    Dim depth As Long
    Dim startPos As Long
    Dim ch As String
    Dim dummy As String
    n = Len(s)
    startPos = i
    depth = 0
    Do While i <= n
        ch = Mid$(s, i, 1)
        If ch = """" Then
            If Not ReadJsonStringRaw(s, i, dummy, i) Then Exit Function
        ElseIf ch = "{" Or ch = "[" Then
            depth = depth + 1
            i = i + 1
        ElseIf ch = "}" Or ch = "]" Then
            depth = depth - 1
            If depth = 0 Then
                raw = Mid$(s, startPos, i - startPos + 1)
                iNext = i + 1
                ReadJsonNestedRaw = True
                Exit Function
            End If
            i = i + 1
        Else
            i = i + 1
        End If
    Loop
End Function

' Tokenize one JSON object line into parallel arrays (1-based; FIRST key
' occurrence wins). Returns False on any structural failure.
Private Function ParseTopLevelPairs(ByVal line As String, _
                                    ByRef keys() As String, _
                                    ByRef typs() As String, _
                                    ByRef vals() As String, _
                                    ByRef pairCount As Long) As Boolean
    Dim s As String
    Dim n As Long
    Dim i As Long
    Dim j As Long
    Dim keyRaw As String
    Dim valRaw As String
    Dim keyName As String
    Dim ch As String
    Dim vType As String
    Dim vVal As String
    Dim numStart As Long
    Dim dup As Boolean

    pairCount = 0
    s = TrimWs(line)
    n = Len(s)
    If n < 2 Then Exit Function
    If Left$(s, 1) <> "{" Or Right$(s, 1) <> "}" Then Exit Function

    ReDim keys(1 To 16)
    ReDim typs(1 To 16)
    ReDim vals(1 To 16)

    i = SkipWs(s, 2)
    If i = n And Mid$(s, i, 1) = "}" Then
        ParseTopLevelPairs = True   ' empty object
        Exit Function
    End If

    Do
        i = SkipWs(s, i)
        If i > n Then Exit Function
        If Mid$(s, i, 1) <> """" Then Exit Function
        If Not ReadJsonStringRaw(s, i, keyRaw, i) Then Exit Function
        keyName = JsonUnescapeString(keyRaw)
        i = SkipWs(s, i)
        If i > n Then Exit Function
        If Mid$(s, i, 1) <> ":" Then Exit Function
        i = SkipWs(s, i + 1)
        If i > n Then Exit Function

        ch = Mid$(s, i, 1)
        If ch = """" Then
            If Not ReadJsonStringRaw(s, i, valRaw, i) Then Exit Function
            vType = "s"
            vVal = JsonUnescapeString(valRaw)
        ElseIf Mid$(s, i, 4) = "true" Then
            vType = "b"
            vVal = "true"
            i = i + 4
        ElseIf Mid$(s, i, 5) = "false" Then
            vType = "b"
            vVal = "false"
            i = i + 5
        ElseIf Mid$(s, i, 4) = "null" Then
            vType = "z"
            vVal = "null"
            i = i + 4
        ElseIf ch = "{" Or ch = "[" Then
            If Not ReadJsonNestedRaw(s, i, valRaw, i) Then Exit Function
            vType = "c"
            vVal = valRaw
        ElseIf InStr(1, "-+.eE0123456789", ch, vbBinaryCompare) > 0 Then
            numStart = i
            Do While i <= n
                If InStr(1, "-+.eE0123456789", Mid$(s, i, 1), vbBinaryCompare) > 0 Then
                    i = i + 1
                Else
                    Exit Do
                End If
            Loop
            vType = "n"
            vVal = Mid$(s, numStart, i - numStart)
        Else
            Exit Function
        End If

        ' First occurrence wins
        dup = False
        For j = 1 To pairCount
            If keys(j) = keyName Then
                dup = True
                Exit For
            End If
        Next j
        If Not dup Then
            pairCount = pairCount + 1
            If pairCount > UBound(keys) Then
                ReDim Preserve keys(1 To UBound(keys) * 2)
                ReDim Preserve typs(1 To UBound(typs) * 2)
                ReDim Preserve vals(1 To UBound(vals) * 2)
            End If
            keys(pairCount) = keyName
            typs(pairCount) = vType
            vals(pairCount) = vVal
        End If

        i = SkipWs(s, i)
        If i > n Then Exit Function
        ch = Mid$(s, i, 1)
        If ch = "," Then
            i = i + 1
        ElseIf ch = "}" Then
            If i = n Then ParseTopLevelPairs = True   ' must be the LAST char
            Exit Function
        Else
            Exit Function
        End If
    Loop
End Function

' Index of key in the parsed pairs IF its value has the wanted type; 0 when
' absent or wrong-typed. The first occurrence decides (no fall-through).
Private Function PairIndex(ByRef keys() As String, ByRef typs() As String, _
                           ByVal pairCount As Long, ByVal key As String, _
                           ByVal wantType As String) As Long
    Dim j As Long
    For j = 1 To pairCount
        If keys(j) = key Then
            If typs(j) = wantType Then PairIndex = j
            Exit Function
        End If
    Next j
End Function

' Parse one JSONL line into the bookmark-only schema. Public so the self-test
' harness can replay the golden parser vectors against it. bookmark_id and
' change_type must be present AS STRINGS or the parse fails; on failure every
' ByRef output is reset (mirroring the Python twin).
Public Function ParseJsonLine(ByVal line As String, _
                              ByRef bookmarkId As String, _
                              ByRef changeType As String, _
                              ByRef oldText As String, _
                              ByRef newText As String, _
                              ByRef addComment As String, _
                              ByRef applyChange As Variant, _
                              ByRef confidence As String) As Boolean
    Dim keys() As String
    Dim typs() As String
    Dim vals() As String
    Dim pairCount As Long
    Dim idx As Long

    On Error GoTo ErrFail

    bookmarkId = ""
    changeType = ""
    oldText = ""
    newText = ""
    addComment = ""
    applyChange = Empty
    confidence = ""

    If Not ParseTopLevelPairs(line, keys, typs, vals, pairCount) Then GoTo ErrFail

    idx = PairIndex(keys, typs, pairCount, "bookmark_id", "s")
    If idx = 0 Then GoTo ErrFail
    bookmarkId = vals(idx)

    idx = PairIndex(keys, typs, pairCount, "change_type", "s")
    If idx = 0 Then GoTo ErrFail
    changeType = vals(idx)

    idx = PairIndex(keys, typs, pairCount, "old_text", "s")
    If idx > 0 Then oldText = vals(idx)

    idx = PairIndex(keys, typs, pairCount, "new_text", "s")
    If idx > 0 Then newText = vals(idx)

    idx = PairIndex(keys, typs, pairCount, "add_comment", "s")
    If idx > 0 Then addComment = vals(idx)

    idx = PairIndex(keys, typs, pairCount, "apply_change", "b")
    If idx > 0 Then applyChange = (vals(idx) = "true")

    idx = PairIndex(keys, typs, pairCount, "confidence", "s")
    If idx > 0 Then confidence = vals(idx)

    ParseJsonLine = True
    Exit Function

ErrFail:
    bookmarkId = ""
    changeType = ""
    oldText = ""
    newText = ""
    addComment = ""
    applyChange = Empty
    confidence = ""
    ParseJsonLine = False
End Function

' Validate a parsed change against the serializer contract. Returns "" when
' valid, else a stable reason code shared with the Python twin and the golden
' vectors. The check ORDER is part of the contract -- mirror exactly.
Public Function ValidateParsedChange(ByVal bookmarkId As String, _
                                     ByVal changeType As String, _
                                     ByVal oldText As String, _
                                     ByVal newText As String, _
                                     ByVal addComment As String) As String
    Dim b As String
    Dim ct As String
    Dim isCommentTarget As Boolean

    b = TrimWs(bookmarkId)
    ct = LCase$(TrimWs(changeType))

    If Len(b) = 0 Then
        ValidateParsedChange = "MISSING_BOOKMARK"
        Exit Function
    End If
    If Len(ct) = 0 Then
        ValidateParsedChange = "MISSING_CHANGE_TYPE"
        Exit Function
    End If

    Select Case ct
        Case "replace_text", "delete_element", "add_comment_only", _
             "reply_to_comment", "accept_revision", "reject_revision"
            ' known type
        Case Else
            ValidateParsedChange = "UNKNOWN_CHANGE_TYPE"
            Exit Function
    End Select

    isCommentTarget = (Left$(bookmarkId, 11) = "AR_COMMENT_")

    If ct = "replace_text" Then
        If Len(TrimWs(newText)) = 0 Then
            ValidateParsedChange = "REPLACE_REQUIRES_NEW_TEXT"
            Exit Function
        End If
    End If
    If ct = "add_comment_only" Then
        If Len(TrimWs(addComment)) = 0 Then
            ValidateParsedChange = "COMMENT_REQUIRES_TEXT"
            Exit Function
        End If
    End If
    If ct = "reply_to_comment" Then
        If Not isCommentTarget Then
            ValidateParsedChange = "REPLY_REQUIRES_COMMENT_TARGET"
            Exit Function
        End If
        If Len(TrimWs(addComment)) = 0 Then
            ValidateParsedChange = "REPLY_REQUIRES_TEXT"
            Exit Function
        End If
    End If
    If ct = "accept_revision" Or ct = "reject_revision" Then
        If isCommentTarget Then
            ValidateParsedChange = "REVISION_REQUIRES_RANGE_TARGET"
            Exit Function
        End If
    End If
    If ct = "add_footnote" Then
        If isCommentTarget Then
            ValidateParsedChange = "FOOTNOTE_REQUIRES_RANGE_TARGET"
            Exit Function
        End If
        If Len(TrimWs(newText)) = 0 Then
            ValidateParsedChange = "FOOTNOTE_REQUIRES_TEXT"
            Exit Function
        End If
    End If

    ValidateParsedChange = ""
End Function

' Map a validation code to the human-readable Log/skip message.
Private Function ValidationMessage(ByVal code As String, ByVal changeType As String) As String
    Select Case code
        Case "MISSING_BOOKMARK"
            ValidationMessage = "Missing bookmark_id"
        Case "MISSING_CHANGE_TYPE"
            ValidationMessage = "Missing change_type"
        Case "UNKNOWN_CHANGE_TYPE"
            ValidationMessage = "Unknown change_type: " & changeType
        Case "REPLACE_REQUIRES_NEW_TEXT"
            ValidationMessage = "replace_text requires non-empty new_text"
        Case "COMMENT_REQUIRES_TEXT"
            ValidationMessage = "add_comment_only requires add_comment text"
        Case "REPLY_REQUIRES_COMMENT_TARGET"
            ValidationMessage = "reply_to_comment requires a comment target (AR_COMMENT_#)"
        Case "REPLY_REQUIRES_TEXT"
            ValidationMessage = "reply_to_comment requires add_comment text"
        Case "REVISION_REQUIRES_RANGE_TARGET"
            ValidationMessage = "accept/reject_revision requires a bookmark range target, not a comment"
        Case "FOOTNOTE_REQUIRES_TEXT"
            ValidationMessage = "add_footnote requires the citation body in new_text"
        Case "FOOTNOTE_REQUIRES_RANGE_TARGET"
            ValidationMessage = "add_footnote requires a bookmark range target, not a comment"
        Case Else
            ValidationMessage = code
    End Select
End Function

'=== Comment coverage (twin: ref/coverage.py) ========================
' A comment is "addressed" iff some reply_to_comment or add_comment_only edit
' targets its AR_COMMENT_ id. Returns the ordered, comma-joined list of the
' commentIds that were NOT addressed ("" if all addressed). The apply step
' warn-gates on a non-empty result; the export persists commentIds.
Public Function ComputeUnaddressed(ByRef commentIds() As String, ByVal idCount As Long, _
                                   ByRef changeTypes() As String, ByRef bookmarks() As String, _
                                   ByVal editCount As Long) As String
    Dim addressed As Object
    Dim i As Long
    Dim ct As String
    Dim b As String
    Dim res As String

    Set addressed = CreateObject("Scripting.Dictionary")
    For i = 1 To editCount
        ct = LCase$(TrimWs(changeTypes(i)))
        If ct = "reply_to_comment" Or ct = "add_comment_only" Then
            b = TrimWs(bookmarks(i))
            If Left$(b, 11) = "AR_COMMENT_" Then addressed(b) = True
        End If
    Next i

    res = ""
    For i = 1 To idCount
        b = TrimWs(commentIds(i))
        If Len(b) > 0 Then
            If Not addressed.Exists(b) Then
                If Len(res) > 0 Then res = res & ","
                res = res & b
            End If
        End If
    Next i
    ComputeUnaddressed = res
End Function

'=== Session-binding gate (twin: ref/session.py) =====================
' Bookmark ids are generic ordinals, so a stale payload can apply cleanly to
' the WRONG document. The serializer's first output line must be:
'   {"meta": "autoreviewer", "session": "<token>", "count": N}
' with the export fingerprint carried verbatim. Default-deny on any mismatch.
Public Function CheckSessionGate(ByRef allLines() As String, ByVal n As Long, _
                                 ByVal expectedToken As String, _
                                 ByRef failCode As String) As Boolean
    Dim sessionTok As String
    Dim cnt As Long

    failCode = ""
    If Len(expectedToken) = 0 Then
        failCode = "NO_EXPORT_TOKEN"
        Exit Function
    End If
    If n <= 0 Then
        failCode = "NO_PAYLOAD"
        Exit Function
    End If
    If Not ParseMetaLine(allLines(1), sessionTok, cnt) Then
        failCode = "META_MISSING"
        Exit Function
    End If
    If sessionTok <> expectedToken Then
        failCode = "TOKEN_MISMATCH"
        Exit Function
    End If
    If cnt <> n - 1 Then
        failCode = "COUNT_MISMATCH"
        Exit Function
    End If
    CheckSessionGate = True
End Function

' Parse the serializer meta line. count must be a plain integer: an optional
' leading minus, then 1-9 digits (both twins enforce the same rule).
Private Function ParseMetaLine(ByVal line As String, _
                               ByRef sessionTok As String, _
                               ByRef cnt As Long) As Boolean
    Dim keys() As String
    Dim typs() As String
    Dim vals() As String
    Dim pairCount As Long
    Dim idx As Long
    Dim rawCnt As String
    Dim body As String
    Dim j As Long

    If Not ParseTopLevelPairs(line, keys, typs, vals, pairCount) Then Exit Function

    idx = PairIndex(keys, typs, pairCount, "meta", "s")
    If idx = 0 Then Exit Function
    If vals(idx) <> "autoreviewer" Then Exit Function

    idx = PairIndex(keys, typs, pairCount, "session", "s")
    If idx = 0 Then Exit Function
    sessionTok = vals(idx)

    idx = PairIndex(keys, typs, pairCount, "count", "n")
    If idx = 0 Then Exit Function
    rawCnt = vals(idx)

    If Left$(rawCnt, 1) = "-" Then body = Mid$(rawCnt, 2) Else body = rawCnt
    If Len(body) = 0 Or Len(body) > 9 Then Exit Function
    For j = 1 To Len(body)
        If InStr(1, "0123456789", Mid$(body, j, 1), vbBinaryCompare) = 0 Then Exit Function
    Next j
    cnt = CLng(rawCnt)
    If cnt < 0 Then Exit Function

    ParseMetaLine = True
End Function

' Map a session-gate code to operator guidance for the abort MsgBox.
Private Function SessionFailMessage(ByVal code As String) As String
    Select Case code
        Case "NO_EXPORT_TOKEN"
            SessionFailMessage = "No export fingerprint is recorded in Config. Run the export step first; the serializer needs its session token."
        Case "NO_PAYLOAD"
            SessionFailMessage = "No payload lines were found."
        Case "META_MISSING"
            SessionFailMessage = "The first line is not a valid AutoReviewer meta line. Expected: {""meta"": ""autoreviewer"", ""session"": ""<token>"", ""count"": N}. Re-run Hand off to Serializer and paste its FULL output."
        Case "TOKEN_MISMATCH"
            SessionFailMessage = "The payload's session token does not match the last export. This JSONL belongs to a DIFFERENT document or run; applying it could write edits into the wrong file."
        Case "COUNT_MISMATCH"
            SessionFailMessage = "The meta line's count does not match the number of edit lines pasted. Lines may be missing or duplicated; paste the serializer's full output again."
        Case Else
            SessionFailMessage = code
    End Select
End Function

' Twin: json_unescape in ref/jsonl_contract.py. Each case advances i itself
' (the simple escapes by 2, a \uXXXX by 6), so \u can consume its 4 hex digits.
' A \uXXXX decodes to the UTF-16 code unit 0xXXXX via ChrW; adjacent surrogate
' halves therefore form the character naturally in VBA's UTF-16 string, so no
' explicit pair-combining is needed here. A malformed \u (not 4 hex digits) is
' an unknown escape: drop the backslash, keep the 'u'.
Private Function JsonUnescapeString(ByVal s As String) As String
    Dim i As Long
    Dim n As Long
    Dim ch As String
    Dim esc As String
    Dim result As String

    n = Len(s)
    i = 1
    Do While i <= n
        ch = Mid$(s, i, 1)
        If ch = "\" And i < n Then
            esc = Mid$(s, i + 1, 1)
            Select Case esc
                Case "\"    ' backslash
                    result = result & "\"
                    i = i + 2
                Case """"   ' double quote
                    result = result & """"
                    i = i + 2
                Case "/"    ' forward slash
                    result = result & "/"
                    i = i + 2
                Case "b"    ' backspace
                    result = result & Chr$(8)
                    i = i + 2
                Case "f"    ' form feed
                    result = result & Chr$(12)
                    i = i + 2
                Case "n"    ' newline
                    result = result & vbLf
                    i = i + 2
                Case "r"    ' carriage return
                    result = result & vbCr
                    i = i + 2
                Case "t"    ' tab
                    result = result & vbTab
                    i = i + 2
                Case "u"
                    If IsHex4(s, i + 2) Then
                        result = result & ChrW(CLng("&H" & Mid$(s, i + 2, 4) & "&"))
                        i = i + 6
                    Else
                        ' Malformed \u: drop backslash, keep the 'u'
                        result = result & esc
                        i = i + 2
                    End If
                Case Else
                    ' Unknown escape: keep literal escaped char
                    result = result & esc
                    i = i + 2
            End Select
        Else
            result = result & ch
            i = i + 1
        End If
    Loop

    JsonUnescapeString = result
End Function

' True if s has 4 hex digits starting at position p.
Private Function IsHex4(ByVal s As String, ByVal p As Long) As Boolean
    Dim k As Long
    Dim c As String
    If p + 3 > Len(s) Then Exit Function
    For k = 0 To 3
        c = Mid$(s, p + k, 1)
        If InStr(1, "0123456789abcdefABCDEF", c, vbBinaryCompare) = 0 Then Exit Function
    Next k
    IsHex4 = True
End Function

' Confidence normalization (unchanged)
Private Function NormalizeConfidence(ByVal raw As String, _
                                     ByVal defaultLevel As String) As String
    Dim s As String
    Dim def As String
    
    ' Normalize defaultLevel first
    def = LCase$(Trim$(defaultLevel))
    Select Case def
        Case "high"
            def = "High"
        Case "low"
            def = "Low"
        Case Else
            def = "Medium"
    End Select
    
    s = LCase$(Trim$(raw))
    
    Select Case s
        Case "high", "h"
            NormalizeConfidence = "High"
        Case "medium", "med", "m"
            NormalizeConfidence = "Medium"
        Case "low", "l"
            NormalizeConfidence = "Low"
        Case Else
            NormalizeConfidence = def
    End Select
End Function

' Add a comment, always authored "AutoReviewer" so AI output is attributable
' (comment author is settable via the object model even when revision author is
' account-driven). Anchor ids are scrubbed from the body; confidence may be
' shown as a text prefix.
Private Sub AddArComment(ByVal wdDoc As Object, _
                          ByVal rng As Object, _
                          ByVal commentText As String, _
                          ByVal confidence As String, _
                          ByVal usePrefix As Boolean)
    Dim finalText As String
    Dim cmNew As Object

    On Error Resume Next
    commentText = modSysUtils.StripArTokens(commentText)
    If Len(Trim$(commentText)) = 0 Then Exit Sub

    finalText = commentText
    If usePrefix Then finalText = "[AR " & confidence & "] " & commentText

    Set cmNew = wdDoc.Comments.Add(Range:=rng, Text:=finalText)
    If Not cmNew Is Nothing Then cmNew.Author = "AutoReviewer"
    On Error GoTo 0
End Sub

Private Sub AddArCommentReply(ByVal wdDoc As Object, _
                               ByVal parentComment As Object, _
                               ByVal commentText As String, _
                               ByVal confidence As String, _
                               ByVal usePrefix As Boolean)
    Dim finalText As String
    Dim cmNew As Object

    On Error Resume Next
    commentText = modSysUtils.StripArTokens(commentText)
    If Len(Trim$(commentText)) = 0 Then Exit Sub

    finalText = commentText
    If usePrefix Then finalText = "[AR " & confidence & "] " & commentText

    Set cmNew = parentComment.Replies.Add(Range:=parentComment.Scope, Text:=finalText)
    If cmNew Is Nothing Then
        ' Fallback to adding a new comment at the same scope
        Set cmNew = wdDoc.Comments.Add(Range:=parentComment.Scope, Text:=finalText)
    End If

    If Not cmNew Is Nothing Then cmNew.Author = "AutoReviewer"
    On Error GoTo 0
End Sub

' Surgical write-back: replace only the differing middle between the range's
' existing text and newText, so a one-word change is a one-word tracked revision
' rather than a whole-paragraph delete+insert. Common prefix/suffix are computed
' on the punctuation-normalized text (1:1, so offsets map onto the real range),
' which also means a smart-quote/dash already in the document is treated as equal
' to a straight one from the model and is left untouched. Returns True if a
' change was written, False if the text is identical.
Private Function ReplaceMinimalSpan(ByVal rng As Object, ByVal newText As String) As Boolean
    Dim existing As String
    Dim na As String, nb As String
    Dim p As Long, s As Long, maxS As Long
    Dim origStart As Long, origEnd As Long
    Dim newMiddle As String

    existing = CStr(rng.Text)
    na = modSysUtils.NormalizePunctuation(existing)
    nb = modSysUtils.NormalizePunctuation(newText)

    p = CommonPrefixLen(na, nb)

    maxS = Len(na)
    If Len(nb) < maxS Then maxS = Len(nb)
    maxS = maxS - p
    If maxS < 0 Then maxS = 0
    s = CommonSuffixLen(na, nb, maxS)

    ' Identical under normalization -> nothing to change.
    If (p + s) >= Len(na) And (p + s) >= Len(nb) Then
        ReplaceMinimalSpan = False
        Exit Function
    End If

    origStart = rng.Start
    origEnd = rng.End
    newMiddle = Mid$(newText, p + 1, Len(newText) - p - s)

    rng.End = origEnd - s
    rng.Start = origStart + p
    rng.Text = newMiddle
    ReplaceMinimalSpan = True
End Function

' Length of the longest common prefix of a and b.
Private Function CommonPrefixLen(ByVal a As String, ByVal b As String) As Long
    Dim n As Long, i As Long
    n = Len(a)
    If Len(b) < n Then n = Len(b)
    For i = 1 To n
        If Mid$(a, i, 1) <> Mid$(b, i, 1) Then
            CommonPrefixLen = i - 1
            Exit Function
        End If
    Next i
    CommonPrefixLen = n
End Function

' Length of the longest common suffix of a and b, capped at maxLen (so prefix
' and suffix do not overlap).
Private Function CommonSuffixLen(ByVal a As String, ByVal b As String, ByVal maxLen As Long) As Long
    Dim i As Long
    For i = 1 To maxLen
        If Mid$(a, Len(a) - i + 1, 1) <> Mid$(b, Len(b) - i + 1, 1) Then
            CommonSuffixLen = i - 1
            Exit Function
        End If
    Next i
    CommonSuffixLen = maxLen
End Function


