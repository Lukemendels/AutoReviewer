Attribute VB_Name = "modWordStamping"
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

