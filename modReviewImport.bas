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
    
    ' Config-driven behavior
    Dim defaultConfidenceLevel As String
    Dim useArPrefix As Boolean
    Dim useArAuthorNames As Boolean
    
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
    useArAuthorNames = GetConfigBool("UseArAuthorNames", False)
    
    '---------------------------
    ' 2) Get JSONL from LLM_Changes (A8:A…)
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
    
    ReDim tmpLines(1 To lastRow - 7)
    n = 0
    For r = 8 To lastRow
        tmp = Trim$(CStr(wsChanges.Cells(r, "A").value))
        If Len(tmp) > 0 Then
            n = n + 1
            tmpLines(n) = tmp
        End If
    Next r
    
    If n = 0 Then
        MsgBox "No non-empty JSONL lines found in column A starting at A8.", vbExclamation
        GoTo Cleanup
    End If
    
    ReDim Preserve tmpLines(1 To n)
    lines = tmpLines
    
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
            .Columns("A:H").EntireColumn.AutoFit
        End With
    Else
        ' Ensure headers are aligned with bookmark_id schema
        If CStr(wsLog.Range("C1").value) <> "BookmarkId" Then wsLog.Range("C1").value = "BookmarkId"
        If CStr(wsLog.Range("H1").value) <> "Confidence" Then wsLog.Range("H1").value = "Confidence"
        wsLog.Columns("A:H").EntireColumn.AutoFit
    End If
    
    logRow = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).row + 1
    
    '---------------------------
    ' 3) Open Word and document
    '---------------------------
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = True
    On Error Resume Next
    wdApp.DisplayAlerts = 0   ' wdAlertsNone
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
    
    wdDoc.TrackRevisions = True
    
    '---------------------------
    ' 4) Process each JSON line (bookmark-only)
    '---------------------------
    totalLines = UBound(lines) - LBound(lines) + 1
    Application.StatusBar = "Applying bookmark-based suggestions: 0 of " & totalLines
    
    For i = LBound(lines) To UBound(lines)
        Dim logStatus As String
        Dim logReason As String
        
        line = Trim$(lines(i))
        logStatus = ""
        logReason = ""
        Dim commentTarget As Object
        applyThis = True
        Set targetRange = Nothing
        Set commentTarget = Nothing
        
        Application.StatusBar = "Applying bookmark-based suggestions: " & (i - LBound(lines) + 1) & " of " & totalLines
        DoEvents
        
        ' Blank line
        If Len(line) = 0 Then
            logStatus = "Skipped"
            logReason = "Blank line"
            skippedCount = skippedCount + 1
            GoTo LogAndNext
        End If
        
        ' Reset per-line variables
        bookmarkId = ""
        changeType = ""
        oldText = ""
        newText = ""
        addComment = ""
        applyChange = Empty
        confidenceRaw = ""
        confidenceNorm = defaultConfidenceLevel
        
        ' Parse JSON line into fields (bookmark-only schema)
        If Not ParseJsonLine(line, bookmarkId, changeType, oldText, newText, _
                             addComment, applyChange, confidenceRaw) Then
            logStatus = "Skipped"
            logReason = "ParseJsonLine failed"
            skippedCount = skippedCount + 1
            GoTo LogAndNext
        End If
        
        parsedOk = parsedOk + 1
        
        ' Normalize confidence using config default if needed
        confidenceNorm = NormalizeConfidence(confidenceRaw, defaultConfidenceLevel)
        
        ' Respect apply_change flag (must be boolean true to apply)
        applyThis = True
        If Not IsEmpty(applyChange) Then
            If VarType(applyChange) = vbBoolean Then
                applyThis = applyChange
            ElseIf VarType(applyChange) = vbString Then
                applyThis = (LCase$(CStr(applyChange)) = "true")
            End If
        End If
        
        If Not applyThis Then
            logStatus = "Skipped"
            logReason = "apply_change=false"
            skippedCount = skippedCount + 1
            GoTo LogAndNext
        End If
        
        ' Basic validation
        If Len(Trim$(bookmarkId)) = 0 Then
            logStatus = "Skipped"
            logReason = "Missing bookmark_id"
            skippedCount = skippedCount + 1
            GoTo LogAndNext
        End If
        
        If Len(Trim$(changeType)) = 0 Then
            logStatus = "Skipped"
            logReason = "Missing change_type"
            skippedCount = skippedCount + 1
            GoTo LogAndNext
        End If
        
        ' Locate bookmark or comment
        If Left$(bookmarkId, 12) = "AR_COMMENT_" Then
            Dim cIndex As Long
            cIndex = Val(Mid$(bookmarkId, 13))
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
        Select Case LCase$(changeType)
            Case CHANGE_REPLACE_TEXT
                If Len(Trim$(newText)) = 0 Then
                    logStatus = "Skipped"
                    logReason = "replace_text requires non-empty new_text"
                    skippedCount = skippedCount + 1
                    GoTo LogAndNext
                End If
                
                Dim editRange As Object
                Dim txtRange As String
                
                ' Work on a copy of the bookmark range
                Set editRange = targetRange.Duplicate
                
                ' For paragraph bookmarks, exclude the trailing paragraph mark
                ' so we don't delete the separator between this paragraph and the next.
                If Left$(bookmarkId, 9) = "AR_PARA_" Then
                    txtRange = CStr(editRange.Text)
                    If Len(txtRange) > 0 And Right$(txtRange, 1) = Chr$(13) Then
                        editRange.End = editRange.End - 1   ' trim off the paragraph mark
                    End If
                End If
                
                ' Replace the text within the adjusted range
                editRange.Text = newText
                
                ' Re-add bookmark on the updated content range
                On Error Resume Next
                wdDoc.Bookmarks.Add name:=bookmarkId, Range:=editRange
                On Error GoTo ErrHandler
                
                appliedCount = appliedCount + 1
                logStatus = "Applied"
                logReason = "replace_text"
                
                If Len(addComment) > 0 Then
                    AddArComment wdDoc, editRange, addComment, confidenceNorm, _
                                  useArPrefix, useArAuthorNames
                End If
           
            Case CHANGE_DELETE_ELEMENT
                ' Delete entire range
                targetRange.Text = ""
                
                ' Re-add a zero-length bookmark at this location so the ID remains valid
                On Error Resume Next
                wdDoc.Bookmarks.Add name:=bookmarkId, Range:=targetRange
                On Error GoTo ErrHandler
                
                appliedCount = appliedCount + 1
                logStatus = "Applied"
                logReason = "delete_element"
                
                If Len(addComment) > 0 Then
                    AddArComment wdDoc, targetRange, addComment, confidenceNorm, _
                                  useArPrefix, useArAuthorNames
                End If
            
            Case CHANGE_ADD_COMMENT
                If Len(Trim$(addComment)) = 0 Then
                    logStatus = "Skipped"
                    logReason = "add_comment_only requires add_comment text"
                    skippedCount = skippedCount + 1
                    GoTo LogAndNext
                End If
                
                AddArComment wdDoc, targetRange, addComment, confidenceNorm, _
                              useArPrefix, useArAuthorNames
                appliedCount = appliedCount + 1
                logStatus = "Applied"
                logReason = "add_comment_only"
                
            Case CHANGE_REPLY_COMMENT
                If Not commentTarget Is Nothing Then
                    If Len(Trim$(addComment)) = 0 Then
                        logStatus = "Skipped"
                        logReason = "reply_to_comment requires add_comment text"
                        skippedCount = skippedCount + 1
                        GoTo LogAndNext
                    End If
                    
                    AddArCommentReply wdDoc, commentTarget, addComment, confidenceNorm, _
                                       useArPrefix, useArAuthorNames
                    
                    appliedCount = appliedCount + 1
                    logStatus = "Applied"
                    logReason = "reply_to_comment"
                Else
                    logStatus = "Skipped"
                    logReason = "reply_to_comment requires a comment target (AR_COMMENT_#)"
                    skippedCount = skippedCount + 1
                    GoTo LogAndNext
                End If
                
            Case CHANGE_ACCEPT_REVISION
                If targetRange Is Nothing Then
                    logStatus = "Skipped"
                    logReason = "accept_revision requires a bookmark target"
                    skippedCount = skippedCount + 1
                    GoTo LogAndNext
                End If
                
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
                If targetRange Is Nothing Then
                    logStatus = "Skipped"
                    logReason = "reject_revision requires a bookmark target"
                    skippedCount = skippedCount + 1
                    GoTo LogAndNext
                End If
                
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
        
        ' Write log entry (includes bookmark_id and Confidence)
        If Not wsLog Is Nothing Then
            On Error Resume Next
            wsLog.Cells(logRow, 1).value = Now
            wsLog.Cells(logRow, 2).value = i - LBound(lines) + 1   ' JSONL line number
            wsLog.Cells(logRow, 3).value = bookmarkId
            wsLog.Cells(logRow, 4).value = changeType
            wsLog.Cells(logRow, 5).value = logStatus
            wsLog.Cells(logRow, 6).value = logReason
            wsLog.Cells(logRow, 7).value = line
            wsLog.Cells(logRow, 8).value = confidenceNorm
            logRow = logRow + 1
            On Error GoTo ErrHandler
        End If
    Next i
    
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
    
    '---------------------------
    ' 6) Sequential Teardown
    ' Close Word safely BEFORE showing Excel MsgBoxes
    '---------------------------
    On Error Resume Next
    Application.StatusBar = False
    If Not wdDoc Is Nothing Then wdDoc.Close SaveChanges:=True
    If Not wdApp Is Nothing Then
        wdApp.NormalTemplate.Saved = True
        wdApp.Quit SaveChanges:=False
    End If
    Set wdDoc = Nothing
    Set wdApp = Nothing
    DoEvents
    On Error GoTo ErrHandler
    
    ' 7) Show Final Summary
    MsgBox "JSONL lines: " & totalLines & vbCrLf & _
           "Parsed OK: " & parsedOk & vbCrLf & _
           "Applied: " & appliedCount & vbCrLf & _
           "Skipped: " & skippedCount, _
           vbInformation, "Apply Bookmark-Based Suggestions"
    
    Exit Sub

Cleanup:
    On Error Resume Next
    Application.StatusBar = False
    If Not wdDoc Is Nothing Then wdDoc.Close SaveChanges:=True
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
    MsgBox "Error applying suggestions: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

' Parse one JSONL line into the bookmark-only schema
Private Function ParseJsonLine(ByVal line As String, _
                               ByRef bookmarkId As String, _
                               ByRef changeType As String, _
                               ByRef oldText As String, _
                               ByRef newText As String, _
                               ByRef addComment As String, _
                               ByRef applyChange As Variant, _
                               ByRef confidence As String) As Boolean
    Dim sVal As String
    Dim bVal As Boolean
    
    On Error GoTo ErrFail
    
    line = Trim$(line)
    If Len(line) = 0 Then Exit Function
    
    If Left$(line, 1) <> "{" Or Right$(line, 1) <> "}" Then
        GoTo ErrFail
    End If
    
    ' bookmark_id (required)
    bookmarkId = ""
    If Not ExtractJsonString(line, "bookmark_id", sVal) Then GoTo ErrFail
    bookmarkId = sVal
    
    ' change_type (required)
    changeType = ""
    If Not ExtractJsonString(line, "change_type", sVal) Then GoTo ErrFail
    changeType = sVal
    
    ' old_text (optional)
    oldText = ""
    If ExtractJsonString(line, "old_text", sVal) Then
        oldText = sVal
    End If
    
    ' new_text (optional, but required for replace_text)
    newText = ""
    If ExtractJsonString(line, "new_text", sVal) Then
        newText = sVal
    End If
    
    ' add_comment (optional)
    addComment = ""
    If ExtractJsonString(line, "add_comment", sVal) Then
        addComment = sVal
    End If
    
    ' apply_change (required in schema but we treat missing as True)
    applyChange = Empty
    If ExtractJsonBoolean(line, "apply_change", bVal) Then
        applyChange = bVal
    End If
    
    ' confidence (optional string)
    confidence = ""
    If ExtractJsonString(line, "confidence", sVal) Then
        confidence = sVal
    End If
    
    ParseJsonLine = True
    Exit Function
    
ErrFail:
    ParseJsonLine = False
End Function

'=== Generic JSON helpers (unchanged) ===============================

Private Function ExtractJsonString(ByVal line As String, ByVal key As String, ByRef value As String) As Boolean
    Dim pattern As String
    Dim posKey As Long, posColon As Long
    Dim i As Long
    Dim sb As String
    Dim ch As String, prevCh As String
    
    pattern = """" & key & """"
    posKey = InStr(1, line, pattern, vbTextCompare)
    If posKey = 0 Then Exit Function
    
    posColon = InStr(posKey + Len(pattern), line, ":", vbTextCompare)
    If posColon = 0 Then Exit Function
    
    ' Move to first non-space after colon
    i = posColon + 1
    Do While i <= Len(line) And Mid$(line, i, 1) = " "
        i = i + 1
    Loop
    
    If i > Len(line) Or Mid$(line, i, 1) <> """" Then Exit Function
    
    ' Start after the opening quote
    i = i + 1
    sb = ""
    prevCh = ""
    
    Do While i <= Len(line)
        ch = Mid$(line, i, 1)
        If ch = """" And prevCh <> "\" Then
            Exit Do              ' reached the closing quote, do NOT include it
        End If
        sb = sb & ch
        prevCh = ch
        i = i + 1
    Loop
    
    value = JsonUnescapeString(sb)
    ExtractJsonString = True
End Function

Private Function ExtractJsonBoolean(ByVal line As String, ByVal key As String, ByRef value As Boolean) As Boolean
    Dim pattern As String
    Dim posKey As Long, posColon As Long
    Dim i As Long, startPos As Long
    Dim ch As String
    Dim token As String
    
    pattern = """" & key & """"
    posKey = InStr(1, line, pattern, vbTextCompare)
    If posKey = 0 Then Exit Function
    
    posColon = InStr(posKey + Len(pattern), line, ":", vbTextCompare)
    If posColon = 0 Then Exit Function
    
    i = posColon + 1
    Do While i <= Len(line) And Mid$(line, i, 1) = " "
        i = i + 1
    Loop
    
    If i > Len(line) Then Exit Function
    
    startPos = i
    For i = startPos To Len(line)
        ch = Mid$(line, i, 1)
        If ch = "," Or ch = "}" Or ch = " " Then Exit For
    Next i
    
    token = LCase$(Trim$(Mid$(line, startPos, i - startPos)))
    If token = "true" Then
        value = True
        ExtractJsonBoolean = True
    ElseIf token = "false" Then
        value = False
        ExtractJsonBoolean = True
    End If
End Function

Private Function JsonUnescapeString(ByVal s As String) As String
    Dim i As Long
    Dim ch As String
    Dim esc As String
    Dim result As String
    
    i = 1
    Do While i <= Len(s)
        ch = Mid$(s, i, 1)
        If ch = "\" And i < Len(s) Then
            esc = Mid$(s, i + 1, 1)
            Select Case esc
                Case "\"    ' backslash
                    result = result & "\"
                Case """"   ' double quote
                    result = result & """"
                Case "/"    ' forward slash
                    result = result & "/"
                Case "b"    ' backspace
                    result = result & Chr$(8)
                Case "f"    ' form feed
                    result = result & Chr$(12)
                Case "n"    ' newline
                    result = result & vbLf
                Case "r"    ' carriage return
                    result = result & vbCr
                Case "t"    ' tab
                    result = result & vbTab
                Case Else
                    ' Unknown escape: keep literal escaped char
                    result = result & esc
            End Select
            i = i + 2
        Else
            result = result & ch
            i = i + 1
        End If
    Loop
    
    JsonUnescapeString = result
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

' High-level helper that applies AR prefix and optional author labeling
' based on config and normalized confidence.
Private Sub AddArComment(ByVal wdDoc As Object, _
                          ByVal rng As Object, _
                          ByVal commentText As String, _
                          ByVal confidence As String, _
                          ByVal usePrefix As Boolean, _
                          ByVal useAuthorNames As Boolean)
    Dim finalText As String
    Dim cmNew As Object
    Dim prefix As String
    
    On Error Resume Next
    If Len(commentText) = 0 Then Exit Sub
    
    finalText = commentText
    
    ' Optional [AR {LEVEL}] prefix in the comment text
    If usePrefix Then
        prefix = "[AR " & confidence & "] "
        finalText = prefix & commentText
    End If
    
    ' Add the comment
    Set cmNew = wdDoc.Comments.Add(Range:=rng, Text:=finalText)
    
    ' Optional: set author name based on confidence
    If useAuthorNames And Not cmNew Is Nothing Then
        cmNew.Author = "AR (" & confidence & ")"
    End If
End Sub

Private Sub AddArCommentReply(ByVal wdDoc As Object, _
                               ByVal parentComment As Object, _
                               ByVal commentText As String, _
                               ByVal confidence As String, _
                               ByVal usePrefix As Boolean, _
                               ByVal useAuthorNames As Boolean)
    Dim finalText As String
    Dim cmNew As Object
    Dim prefix As String
    
    On Error Resume Next
    If Len(commentText) = 0 Then Exit Sub
    
    finalText = commentText
    If usePrefix Then
        prefix = "[AR " & confidence & "] "
        finalText = prefix & commentText
    End If
    
    Set cmNew = parentComment.Replies.Add(Range:=parentComment.Scope, Text:=finalText)
    If cmNew Is Nothing Then
        ' Fallback to adding a new comment at the same scope
        Set cmNew = wdDoc.Comments.Add(Range:=parentComment.Scope, Text:=finalText)
    End If
    
    If useAuthorNames And Not cmNew Is Nothing Then
        cmNew.Author = "AR (" & confidence & ")"
    End If
    On Error GoTo 0
End Sub


