Attribute VB_Name = "modWordUtils"
Option Explicit


' Stamp paragraphs, table cells, and footnotes with stable AR_ bookmarks.
' Call this once after opening the Word document (wdDoc) via COM.
Public Sub StampDocWithArBookmarks(ByVal wdDoc As Object)
    Dim wdApp As Object
    Dim origPagination As Boolean
    Dim paginationChanged As Boolean

    ' Pagination/repagination during a long stamping pass costs real time on a
    ' large document; suspend it for the duration and restore unconditionally.
    On Error Resume Next
    Set wdApp = wdDoc.Application
    origPagination = wdApp.Options.Pagination
    wdApp.Options.Pagination = False
    paginationChanged = True
    On Error GoTo 0

    ' Clear any stale AR_ anchors before re-stamping. Paragraph indices shift as
    ' a document is edited, so a leftover AR_PARA_00005 from a prior pass would
    ' point at the wrong paragraph; a clean slate keeps anchoring honest.
    RemoveArBookmarks wdDoc
    StampParagraphBookmarks wdDoc
    StampTableCellBookmarks wdDoc
    StampFootnoteBookmarks wdDoc

    If paginationChanged Then
        On Error Resume Next
        wdApp.Options.Pagination = origPagination
        On Error GoTo 0
    End If
End Sub

' Stamp one AR_REV_NNNNN bookmark over each tracked revision, in revision order.
' The numbering is a pure function of the (unchanged) document, so the export
' and the apply step produce identical AR_REV ids without persisting anything.
Public Sub StampRevisionBookmarks(ByVal wdDoc As Object)
    Dim revIdx As Long
    Dim rev As Object
    On Error Resume Next
    If wdDoc.Revisions.Count = 0 Then Exit Sub
    revIdx = 0
    For Each rev In wdDoc.Revisions
        revIdx = revIdx + 1
        wdDoc.Bookmarks.Add Name:="AR_REV_" & Format$(revIdx, "00000"), _
                            Range:=rev.Range
    Next rev
    On Error GoTo 0
End Sub

' Remove every AR_* bookmark from the document. Used as the terminal step of the
' apply pipeline (leave the delivered doc clean) and defensively before
' re-stamping. Iterates descending because Delete reindexes the collection.
' Delete removes a Bookmark and reindexes wdDoc.Bookmarks, so we cannot
' enumerate-and-delete in one pass. Two-phase: For Each to collect the AR_*
' bookmarks into a VBA Collection (the references stay valid as siblings are
' removed), then Delete each one from that snapshot.
Public Sub RemoveArBookmarks(ByVal wdDoc As Object)
    Dim bm As Object
    Dim nm As String
    Dim bmToDelete As Collection
    Dim delBm As Object

    Set bmToDelete = New Collection

    On Error Resume Next
    For Each bm In wdDoc.Bookmarks
        nm = CStr(bm.name)
        If Left$(nm, 3) = "AR_" Then
            bmToDelete.Add bm
        End If
    Next bm

    For Each delBm In bmToDelete
        delBm.Delete
    Next delBm
    On Error GoTo 0
End Sub

' AR_PARA numbering is a pure function of paragraph POSITION (1-based, in
' document order), so the export and apply passes -- which both call this on
' the same unchanged working copy -- produce identical ids. The counter `i`
' increments for EVERY paragraph, including skipped empty ones, so renumbering
' never shifts. We use For Each rather than indexed Paragraphs access: walking
' Word's Paragraphs collection by index restarts from the start of the document
' each time, making that loop O(n^2) on long documents; For Each is a single
' forward pass.
Private Sub StampParagraphBookmarks(ByVal wdDoc As Object)
    Dim i As Long
    Dim para As Object
    Dim paraRange As Object
    Dim bmName As String

    On Error GoTo ErrHandler

    i = 0
    For Each para In wdDoc.Paragraphs
        i = i + 1
        Set paraRange = para.Range

        ' Skip empty/whitespace-only paragraphs (counter above already advanced)
        If Len(Trim$(paraRange.Text)) > 0 Then
            bmName = "AR_PARA_" & Format$(i, "00000")
            wdDoc.Bookmarks.Add name:=bmName, Range:=paraRange
        End If
    Next para

Cleanup:
    Set paraRange = Nothing
    Set para = Nothing
    Exit Sub
ErrHandler:
    Resume Cleanup
End Sub

Private Sub StampTableCellBookmarks(ByVal wdDoc As Object)
    Dim t As Long, r As Long, c As Long
    Dim tbl As Object
    Dim cellRange As Object
    Dim bmName As String
    
    On Error GoTo ErrHandler

    t = 0
    For Each tbl In wdDoc.Tables
        t = t + 1
        For r = 1 To tbl.Rows.Count
            For c = 1 To tbl.Columns.Count
                Set cellRange = tbl.Cell(r, c).Range
                ' Exclude end-of-cell marker
                cellRange.End = cellRange.End - 1
                If Len(Trim$(cellRange.Text)) > 0 Then
                    bmName = "AR_CELL_" & t & "_" & r & "_" & c
                    wdDoc.Bookmarks.Add name:=bmName, Range:=cellRange
                End If
            Next c
        Next r
    Next tbl
    
Cleanup:
    Set cellRange = Nothing
    Set tbl = Nothing
    Exit Sub
ErrHandler:
    Resume Cleanup
End Sub

Private Sub StampFootnoteBookmarks(ByVal wdDoc As Object)
    Dim i As Long
    Dim fnRange As Object
    Dim bmName As String
    
    On Error GoTo ErrHandler
    
    If wdDoc.Footnotes.Count = 0 Then Exit Sub
    
    For i = 1 To wdDoc.Footnotes.Count
        Set fnRange = wdDoc.Footnotes(i).Range
        If Len(Trim$(fnRange.Text)) > 0 Then
            bmName = "AR_FN_" & Format$(i, "000")
            wdDoc.Bookmarks.Add name:=bmName, Range:=fnRange
        End If
    Next i
    
Cleanup:
    Set fnRange = Nothing
    Exit Sub
ErrHandler:
    Resume Cleanup
End Sub



' Test helper:
' - Lets you pick a Word document.
' - Opens Word, creates a "Final" copy (revisions accepted).
' - Calls StampDocWithArBookmarks on that Final copy.
' - Leaves Word and both documents open so you can inspect bookmarks.
Public Sub TestStampFinalCopyBookmarks()
    Const msoFileDialogFilePicker As Long = 3

    Dim wdApp       As Object   ' Word.Application (late-bound)
    Dim wdDocOrig   As Object   ' Original document (read-only)
    Dim wdDocFinal  As Object   ' Final copy (revisions accepted)
    Dim fd          As Object   ' FileDialog
    Dim wordPath    As String
    Dim msg         As String
    Dim paraCount   As Long
    Dim bmCount     As Long

    On Error GoTo ErrHandler

    '---------------------------
    ' 1) Let user pick the Word document
    '---------------------------
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Select Word document to test AR bookmark stamping"
        .AllowMultiSelect = False
        If Len(modAppCore.GetWorkFolder()) > 0 Then
            .InitialFileName = modAppCore.GetWorkFolder() & "\"
        End If

        If .Show <> -1 Then
            ' User cancelled
            GoTo Cleanup
        End If

        wordPath = .SelectedItems(1)
    End With
    Set fd = Nothing

    wordPath = Trim$(wordPath)
    If Len(wordPath) = 0 Then GoTo Cleanup

    '---------------------------
    ' 2) Start Word and open original doc (read-only)
    '---------------------------
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = True

    On Error Resume Next
    wdApp.DisplayAlerts = 0      ' wdAlertsNone
    On Error GoTo ErrHandler

    Set wdDocOrig = wdApp.Documents.Open(Filename:=wordPath, ReadOnly:=True)
    If wdDocOrig Is Nothing Then
        MsgBox "Word could not open the document:" & vbCrLf & wordPath, vbCritical, _
               "TestStampFinalCopyBookmarks"
        GoTo Cleanup
    End If

    '---------------------------
    ' 3) Create a temporary FINAL copy and accept all revisions
    '---------------------------
    Set wdDocFinal = wdApp.Documents.Add
    wdDocFinal.Range.FormattedText = wdDocOrig.Range.FormattedText

    ' Track & accept revisions in the FINAL copy only
    wdDocFinal.TrackRevisions = True
    wdDocFinal.AcceptAllRevisions
    wdDocFinal.TrackRevisions = False

    '---------------------------
    ' 4) Stamp bookmarks on the FINAL copy
    '---------------------------
    ' Requires modWordStamping with:
    '   Public Sub StampDocWithArBookmarks(ByVal wdDoc As Object)
    StampDocWithArBookmarks wdDocFinal

    '---------------------------
    ' 5) Basic diagnostics (bookmark count, paragraph count)
    '---------------------------
    On Error Resume Next
    paraCount = wdDocFinal.Paragraphs.Count
    bmCount = wdDocFinal.Bookmarks.Count
    On Error GoTo ErrHandler

    msg = "Final copy has been created and stamped with AR bookmarks." & vbCrLf & vbCrLf & _
          "Document: " & wordPath & vbCrLf & _
          "Paragraphs: " & paraCount & vbCrLf & _
          "Total bookmarks (all types): " & bmCount & vbCrLf & vbCrLf & _
          "Word is left open. In Word, use Insert -> Bookmark to inspect" & vbCrLf & _
          "AR_PARA_..., AR_CELL_..., and AR_FN_... entries."

    MsgBox msg, vbInformation, "TestStampFinalCopyBookmarks"

    ' IMPORTANT: We deliberately DO NOT close wdDocFinal, wdDocOrig, or wdApp here.
    ' This is a test helper; you inspect the documents manually and close Word yourself.
    GoTo Done

Cleanup:
    ' In this test helper, we intentionally do NOT force close Word or the docs
    ' if they were opened successfully; only clean up if we failed before opening.
    On Error Resume Next
    Set fd = Nothing
    ' Do NOT close wdDocFinal or wdDocOrig here; leave Word state alone.
    ' Do NOT call wdApp.Quit here; let the user close Word manually.
    Set wdDocFinal = Nothing
    Set wdDocOrig = Nothing
    Set wdApp = Nothing
Done:
    Exit Sub

ErrHandler:
    MsgBox "Error in TestStampFinalCopyBookmarks: " & Err.Description, vbCritical
    Resume Cleanup
End Sub



