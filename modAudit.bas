Attribute VB_Name = "modAudit"
Option Explicit

' The logic_trace surface (MKS Normative Core s7 / s9.4; TSA Profile s9.3).
'
' On this substrate the audit lineage is the product, not the safety tax on
' automation (Profile s1.3). Every review run -- one clipboard round-trip from
' the document to a Kernel and back -- appends one row to the Trace sheet
' capturing the things an automated bus structurally lacks: the operator who
' performed the transport and ratification, the recommended route, and a
' transport fingerprint of the exact JSONL that produced the edits. The
' per-edit Log sheet (written by modReviewImport) records what changed; the
' Trace sheet records the lineage of the run that changed it.

Public Sub EnsureTraceSheet(ByRef wsTrace As Worksheet)
    Dim wb As Workbook
    Set wb = ThisWorkbook

    On Error Resume Next
    Set wsTrace = wb.Worksheets("Trace")
    On Error GoTo 0

    If wsTrace Is Nothing Then
        Set wsTrace = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        wsTrace.name = "Trace"
        With wsTrace
            .Range("A1").value = "Timestamp"
            .Range("B1").value = "Operator"
            .Range("C1").value = "Persona"
            .Range("D1").value = "Mode"
            .Range("E1").value = "SourceDoc"
            .Range("F1").value = "WorkingDoc"
            .Range("G1").value = "RecommendedRoute"
            .Range("H1").value = "ExportFingerprint"
            .Range("I1").value = "JsonlFingerprint"
            .Range("J1").value = "LinesTotal"
            .Range("K1").value = "Applied"
            .Range("L1").value = "Skipped"
            .Range("M1").value = "UnaddressedComments"
            .Range("N1").value = "CoverageDecision"
            .Range("O1").value = "WorkFolder"
            .Range("A1:O1").Font.Bold = True
            .Columns("A:O").EntireColumn.AutoFit
        End With
    Else
        ' Backfill the coverage/work-folder columns on a pre-existing Trace sheet.
        If CStr(wsTrace.Range("M1").value) <> "UnaddressedComments" Then wsTrace.Range("M1").value = "UnaddressedComments"
        If CStr(wsTrace.Range("N1").value) <> "CoverageDecision" Then wsTrace.Range("N1").value = "CoverageDecision"
        If CStr(wsTrace.Range("O1").value) <> "WorkFolder" Then wsTrace.Range("O1").value = "WorkFolder"
    End If
End Sub

' Append one logic_trace row for a review/apply run. The coverage columns are
' optional so the export-time row and older callers stay valid.
Public Sub AppendReviewTrace(ByVal mode As String, _
                             ByVal persona As String, _
                             ByVal sourceDoc As String, _
                             ByVal workingDoc As String, _
                             ByVal recommendedRoute As String, _
                             ByVal exportFingerprint As String, _
                             ByVal jsonlFingerprint As String, _
                             ByVal linesTotal As Long, _
                             ByVal appliedCount As Long, _
                             ByVal skippedCount As Long, _
                             Optional ByVal unaddressedComments As String = "", _
                             Optional ByVal coverageDecision As String = "")
    Dim wsTrace As Worksheet
    Dim r As Long
    Dim workFolder As String

    EnsureTraceSheet wsTrace

    r = wsTrace.Cells(wsTrace.Rows.Count, "A").End(xlUp).row + 1
    If r < 2 Then r = 2

    On Error Resume Next
    workFolder = modAppCore.GetWorkFolder()
    On Error GoTo 0

    On Error Resume Next
    wsTrace.Cells(r, 1).value = Now
    wsTrace.Cells(r, 2).value = Environ$("USERNAME")   ' operator-attested transport/ratifier
    wsTrace.Cells(r, 3).value = persona
    wsTrace.Cells(r, 4).value = mode
    wsTrace.Cells(r, 5).value = sourceDoc
    wsTrace.Cells(r, 6).value = workingDoc
    wsTrace.Cells(r, 7).value = recommendedRoute
    wsTrace.Cells(r, 8).value = exportFingerprint
    wsTrace.Cells(r, 9).value = jsonlFingerprint
    wsTrace.Cells(r, 10).value = linesTotal
    wsTrace.Cells(r, 11).value = appliedCount
    wsTrace.Cells(r, 12).value = skippedCount
    wsTrace.Cells(r, 13).value = unaddressedComments
    wsTrace.Cells(r, 14).value = coverageDecision
    wsTrace.Cells(r, 15).value = workFolder
    wsTrace.Columns("A:O").EntireColumn.AutoFit
    On Error GoTo 0
End Sub
