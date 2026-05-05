Attribute VB_Name = "modWordUtils"
Option Explicit


' Stamp paragraphs, table cells, and footnotes with stable MKS_ bookmarks.
' Call this once after opening the Word document (wdDoc) via COM.
Public Sub StampDocWithMksBookmarks(ByVal wdDoc As Object)
    StampParagraphBookmarks wdDoc
    StampTableCellBookmarks wdDoc
    StampFootnoteBookmarks wdDoc
End Sub

Private Sub StampParagraphBookmarks(ByVal wdDoc As Object)
    Dim i As Long
    Dim paraRange As Object
    Dim bmName As String
    
    On Error GoTo ErrHandler
    
    For i = 1 To wdDoc.Paragraphs.Count
        Set paraRange = wdDoc.Paragraphs(i).Range
        
        ' Skip empty/whitespace-only paragraphs
        If Len(Trim$(paraRange.Text)) > 0 Then
            bmName = "MKS_PARA_" & Format$(i, "00000")
            If Not wdDoc.Bookmarks.Exists(bmName) Then
                wdDoc.Bookmarks.Add name:=bmName, Range:=paraRange
            End If
        End If
    Next i
    
Cleanup:
    Set paraRange = Nothing
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
    
    For t = 1 To wdDoc.Tables.Count
        Set tbl = wdDoc.Tables(t)
        For r = 1 To tbl.Rows.Count
            For c = 1 To tbl.Columns.Count
                Set cellRange = tbl.Cell(r, c).Range
                ' Exclude end-of-cell marker
                cellRange.End = cellRange.End - 1
                If Len(Trim$(cellRange.Text)) > 0 Then
                    bmName = "MKS_CELL_" & t & "_" & r & "_" & c
                    If Not wdDoc.Bookmarks.Exists(bmName) Then
                        wdDoc.Bookmarks.Add name:=bmName, Range:=cellRange
                    End If
                End If
            Next c
        Next r
    Next t
    
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
            bmName = "MKS_FN_" & Format$(i, "000")
            If Not wdDoc.Bookmarks.Exists(bmName) Then
                wdDoc.Bookmarks.Add name:=bmName, Range:=fnRange
            End If
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
' - Calls StampDocWithMksBookmarks on that Final copy.
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
        .Title = "Select Word document to test MKS bookmark stamping"
        .AllowMultiSelect = False
        If Len(ThisWorkbook.Path) > 0 Then
            .InitialFileName = ThisWorkbook.Path & "\"
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
    '   Public Sub StampDocWithMksBookmarks(ByVal wdDoc As Object)
    StampDocWithMksBookmarks wdDocFinal

    '---------------------------
    ' 5) Basic diagnostics (bookmark count, paragraph count)
    '---------------------------
    On Error Resume Next
    paraCount = wdDocFinal.Paragraphs.Count
    bmCount = wdDocFinal.Bookmarks.Count
    On Error GoTo ErrHandler

    msg = "Final copy has been created and stamped with MKS bookmarks." & vbCrLf & vbCrLf & _
          "Document: " & wordPath & vbCrLf & _
          "Paragraphs: " & paraCount & vbCrLf & _
          "Total bookmarks (all types): " & bmCount & vbCrLf & vbCrLf & _
          "Word is left open. In Word, use Insert -> Bookmark to inspect" & vbCrLf & _
          "MKS_PARA_..., MKS_CELL_..., and MKS_FN_... entries."

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



