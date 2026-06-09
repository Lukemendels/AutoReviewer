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
    Dim oldBgSave As Boolean
    
    On Error GoTo ErrHandler
    
    '---------------------------
    ' 1) Pick the Word document
    '---------------------------
    Set fd = Application.FileDialog(3) ' msoFileDialogFilePicker
    With fd
        .Title = "Select Word document to export for LLM"
        .AllowMultiSelect = False
        If Len(ThisWorkbook.Path) > 0 Then
            .InitialFileName = ThisWorkbook.Path & "\"
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
    ' 1a) Config sheet: record WordDocPath via key-based helper
    '---------------------------
    EnsureConfigSheet wsConfig
    SetConfigValue "WordDocPath", wordPath
    
    ' Determine export format (plain | markdown)
    exportFormat = LCase$(Trim$(GetConfigValue("ExportFormat", "plain")))
    If exportFormat <> "markdown" Then
        exportFormat = "plain"
    End If
    
    '---------------------------
    ' 2) Start Word and open original doc (READ/WRITE)
    '---------------------------
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = True
    On Error Resume Next
    wdApp.DisplayAlerts = 0
    On Error GoTo ErrHandler
    
    ' Open original document read/write so we can add bookmarks and SAVE them
    Set wdDoc = wdApp.Documents.Open(Filename:=wordPath, ReadOnly:=False)
    If wdDoc Is Nothing Then
        MsgBox "Word could not open the document: " & wordPath, vbCritical, "ExportWordDocForLLM"
        GoTo Cleanup
    End If
    
    '---------------------------
    ' 2a) STAMP DOC WITH AR BOOKMARKS
    '---------------------------
    ' Stamp original (this is the one we will edit later and SAVE to disk)
    StampDocWithArBookmarks wdDoc
    
    ' Stamp Revisions in the original document if we are in Respond Mode
    If isRespondMode Then
        Dim revIdx As Long
        For revIdx = 1 To wdDoc.Revisions.Count
            On Error Resume Next
            wdDoc.Bookmarks.Add Name:="AR_REV_" & Format(revIdx, "00000"), Range:=wdDoc.Revisions(revIdx).Range
            On Error GoTo ErrHandler
        Next revIdx
    End If
    
    '---------------------------
    ' 2b) SAVE the original so AR_* bookmarks persist in the real document
    '---------------------------
    On Error Resume Next
    oldBgSave = wdApp.Options.BackgroundSave
    wdApp.Options.BackgroundSave = False
    wdDoc.Save
    wdApp.Options.BackgroundSave = oldBgSave
    On Error GoTo ErrHandler
    
    '---------------------------
    ' 3) Extract metadata (comments and revisions) BEFORE accepting revisions
    '---------------------------
    buffer = ""
    
    Dim bufferComments As String
    Dim bufferRevisions As String
    
    bufferComments = "<<COMMENTS_START>>" & vbCrLf
    
    If wdDoc.Comments.Count = 0 Then
        bufferComments = bufferComments & "(No comments in document)" & vbCrLf
    Else
        For Each c In wdDoc.Comments
            bufferComments = bufferComments & "## AR_COMMENT_" & c.Index & vbCrLf
            
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
    ' 6) Decide where to save .txt (UTF-8) – User's Documents
    '---------------------------
    docsFolder = GetUserDocumentsFolder()
    
    ' Fallbacks if Documents is not available
    If Len(docsFolder) = 0 Then
        If Len(ThisWorkbook.Path) > 0 Then
            docsFolder = ThisWorkbook.Path
        Else
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
    Dim promptText As String
    If isRespondMode Then
        promptText = "I am attaching an exported Word document containing text, bookmark IDs, reviewer comments, and reviewer tracked changes (revisions). The text already reflects any accepted tracked changes. Please conversationally unpack what the reviewer is asking for in both their comments and their tracked changes. Decide as a team how to edit the document to best address the feedback, and then output a JSON list of edits. You can use 'replace_text' to modify the text, 'reply_to_comment' to respond directly to the reviewer's comments, 'accept_revision' to accept a specific tracked change, and 'reject_revision' to reject a specific tracked change."
    Else
        promptText = "I am attaching an exported Word document containing text and bookmark IDs. Please review the text according to the established style rules, and output a JSON list of edits targeting the specific bookmark IDs."
    End If
    modSysUtils.CopyToClipboard promptText
    
    Dim gptUrl As String
    gptUrl = "https://chat.dhs.gov/workspaces/4cf75bdf-de55-4f01-8c3f-0444ace52010"
    On Error Resume Next
    Dim activePersona As String
    activePersona = modAppCore.GetConfigValue("ActivePersona")
    Dim configUrl As String
    If Len(activePersona) > 0 Then configUrl = modAppCore.GetAssistantUrl(activePersona)
    If Len(configUrl) > 0 Then gptUrl = configUrl
    On Error GoTo 0
    modSysUtils.OpenURL gptUrl
    
    ' 10) Final Excel MsgBox (Word is already closed)
    MsgBox "Export complete (UTF-8):" & vbCrLf & savePath & vbCrLf & vbCrLf & _
           "A prompt has been copied to your clipboard, and your custom GPT URL has been opened." & vbCrLf & _
           "1. Paste the prompt (Ctrl+V) into the GPT." & vbCrLf & _
           "2. Upload/Drop the exported .txt file.", vbInformation, "Export Complete"
    
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

'=== Helpers ========================================================
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
        docText = docForExport.Content.Text
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
        docText = docForExport.Content.Text
    Else
        docText = ""
    End If
    On Error GoTo 0
    
    BuildDocumentTextSection = "<<DOCUMENT_TEXT_START>>" & vbCrLf & _
                               docText & _
                               "<<DOCUMENT_TEXT_END>>" & vbCrLf & vbCrLf
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

