Attribute VB_Name = "modTrainingCorpusBuilder"
Option Explicit

Public Sub AddDocToCorpus()
    Dim wdApp As Object
    Dim wdDoc As Object
    Dim docPath As String
    Dim targetAuthor As String
    Dim fd As FileDialog
    
    Dim activePersona As String
    activePersona = ConfigHelpers.GetConfigValue("ActivePersona")
    
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
    
    targetAuthor = modAuthorFilter.SelectTargetAuthor(wdDoc)
    
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
    modWordStamping.StampDocWithMksBookmarks wdDoc
    
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
    modPersonaRegistry.UpsertPersona activePersona, corpusPath:=corpusPath, incrementTrainingCount:=True
    
    wdDoc.Close False
    MsgBox "Document added to corpus successfully.", vbInformation
End Sub
