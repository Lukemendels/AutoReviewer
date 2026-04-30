Attribute VB_Name = "modBookmarkTest"
Option Explicit

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


