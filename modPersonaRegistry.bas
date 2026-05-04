Attribute VB_Name = "modPersonaRegistry"
Option Explicit

' Ensures the Personas sheet exists and has headers
Public Sub EnsurePersonasSheet(ByRef wsPersonas As Worksheet)
    Dim wb As Workbook
    Dim created As Boolean
    
    Set wb = ThisWorkbook
    
    On Error Resume Next
    Set wsPersonas = wb.Worksheets("Personas")
    On Error GoTo 0
    
    If wsPersonas Is Nothing Then
        created = True
        Set wsPersonas = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        wsPersonas.name = "Personas"
        
        ' Headers
        wsPersonas.Range("A1").value = "PersonaName"
        wsPersonas.Range("B1").value = "AssistantUrl"
        wsPersonas.Range("C1").value = "CorpusPath"
        wsPersonas.Range("D1").value = "SkillMdPath"
        wsPersonas.Range("E1").value = "TrainingDocCount"
        wsPersonas.Range("F1").value = "LastUpdated"
        wsPersonas.Range("G1").value = "Notes"
        
        ' Make headers bold
        wsPersonas.Range("A1:G1").Font.Bold = True
        wsPersonas.Columns("A:G").EntireColumn.AutoFit
    End If
End Sub

' Adds or updates a persona in the registry
Public Sub UpsertPersona(ByVal personaName As String, _
                         Optional ByVal assistantUrl As String = "", _
                         Optional ByVal corpusPath As String = "", _
                         Optional ByVal skillMdPath As String = "", _
                         Optional ByVal incrementTrainingCount As Boolean = False, _
                         Optional ByVal notes As String = "")
    Dim wsPersonas As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim foundRow As Long
    
    EnsurePersonasSheet wsPersonas
    
    lastRow = wsPersonas.Cells(wsPersonas.Rows.Count, "A").End(xlUp).row
    foundRow = 0
    
    For r = 2 To lastRow
        If StrComp(Trim(wsPersonas.Cells(r, "A").value), Trim(personaName), vbTextCompare) = 0 Then
            foundRow = r
            Exit For
        End If
    Next r
    
    If foundRow = 0 Then
        foundRow = lastRow + 1
        wsPersonas.Cells(foundRow, "A").value = personaName
        wsPersonas.Cells(foundRow, "E").value = 0 ' Initialize count
    End If
    
    If Len(assistantUrl) > 0 Then wsPersonas.Cells(foundRow, "B").value = assistantUrl
    If Len(corpusPath) > 0 Then wsPersonas.Cells(foundRow, "C").value = corpusPath
    If Len(skillMdPath) > 0 Then wsPersonas.Cells(foundRow, "D").value = skillMdPath
    
    If incrementTrainingCount Then
        wsPersonas.Cells(foundRow, "E").value = Val(wsPersonas.Cells(foundRow, "E").value) + 1
    End If
    
    wsPersonas.Cells(foundRow, "F").value = Now
    If Len(notes) > 0 Then wsPersonas.Cells(foundRow, "G").value = notes
    
    wsPersonas.Columns("A:G").EntireColumn.AutoFit
End Sub

' Gets the Assistant URL for the given persona
Public Function GetAssistantUrl(ByVal personaName As String) As String
    Dim wsPersonas As Worksheet
    Dim lastRow As Long
    Dim r As Long
    
    EnsurePersonasSheet wsPersonas
    
    lastRow = wsPersonas.Cells(wsPersonas.Rows.Count, "A").End(xlUp).row
    
    For r = 2 To lastRow
        If StrComp(Trim(wsPersonas.Cells(r, "A").value), Trim(personaName), vbTextCompare) = 0 Then
            GetAssistantUrl = CStr(wsPersonas.Cells(r, "B").value)
            Exit Function
        End If
    Next r
    
    GetAssistantUrl = ""
End Function

' Gets a list of all persona names
Public Function GetAllPersonaNames() As Collection
    Dim wsPersonas As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim col As New Collection
    
    EnsurePersonasSheet wsPersonas
    
    lastRow = wsPersonas.Cells(wsPersonas.Rows.Count, "A").End(xlUp).row
    
    For r = 2 To lastRow
        If Len(Trim(wsPersonas.Cells(r, "A").value)) > 0 Then
            col.Add Trim(wsPersonas.Cells(r, "A").value)
        End If
    Next r
    
    Set GetAllPersonaNames = col
End Function
