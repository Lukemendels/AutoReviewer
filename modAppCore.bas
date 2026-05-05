Attribute VB_Name = "modAppCore"
Option Explicit


' Ensures that the Config sheet exists and, if newly created,
' seeds it with standard keys and default values.
Public Sub EnsureConfigSheet(ByRef wsConfig As Worksheet)
    Dim wb As Workbook
    Dim created As Boolean
    
    Set wb = ThisWorkbook
    
    On Error Resume Next
    Set wsConfig = wb.Worksheets("Config")
    On Error GoTo 0
    
    If wsConfig Is Nothing Then
        created = True
        Set wsConfig = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        wsConfig.name = "Config"
        
        ' Headers
        wsConfig.Range("A1").value = "Setting"
        wsConfig.Range("B1").value = "Value"
        
        ' Core keys used today
        wsConfig.Range("A2").value = "WordDocPath"
        wsConfig.Range("B2").value = ""          ' filled by export macro
        
        wsConfig.Range("A3").value = "LastExportTxtPath"
        wsConfig.Range("B3").value = ""          ' filled by export macro
        
        ' Future / config-driven behavior keys (defaults)
        wsConfig.Range("A4").value = "ExportFormat"
        wsConfig.Range("B4").value = "plain"     ' plain | markdown
        
        wsConfig.Range("A5").value = "UseArCommentPrefix"
        wsConfig.Range("B5").value = "FALSE"     ' TRUE | FALSE
        
        wsConfig.Range("A6").value = "DefaultConfidenceLevel"
        wsConfig.Range("B6").value = "Medium"    ' High | Medium | Low
        
        wsConfig.Range("A7").value = "UseArAuthorNames"
        wsConfig.Range("B7").value = "FALSE"     ' TRUE | FALSE
        
        wsConfig.Range("A8").value = "ActivePersona"
        wsConfig.Range("B8").value = ""
        
        wsConfig.Columns("A:B").EntireColumn.AutoFit
    End If
End Sub

' Reads the value for a given key from Config!A:B.
' Returns defaultVal if key is not found or the value is empty.
Public Function GetConfigValue(ByVal keyName As String, _
                               Optional ByVal defaultVal As String = "") As String
    Dim wsConfig As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim keyCell As String
    Dim val As String
    
    EnsureConfigSheet wsConfig
    
    lastRow = wsConfig.Cells(wsConfig.Rows.Count, "A").End(xlUp).row
    If lastRow < 2 Then
        GetConfigValue = defaultVal
        Exit Function
    End If
    
    For r = 2 To lastRow
        keyCell = Trim$(CStr(wsConfig.Cells(r, "A").value))
        If Len(keyCell) > 0 Then
            If StrComp(keyCell, keyName, vbTextCompare) = 0 Then
                val = CStr(wsConfig.Cells(r, "B").value)
                If Len(val) = 0 Then
                    GetConfigValue = defaultVal
                Else
                    GetConfigValue = val
                End If
                Exit Function
            End If
        End If
    Next r
    
    GetConfigValue = defaultVal
End Function

' Writes/updates the value for a given key in Config!A:B.
' If the key exists, updates its value; otherwise appends a new row.
Public Sub SetConfigValue(ByVal keyName As String, ByVal value As String)
    Dim wsConfig As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim keyCell As String
    
    EnsureConfigSheet wsConfig
    
    lastRow = wsConfig.Cells(wsConfig.Rows.Count, "A").End(xlUp).row
    If lastRow < 2 Then lastRow = 1
    
    ' Try to find existing key
    For r = 2 To lastRow
        keyCell = Trim$(CStr(wsConfig.Cells(r, "A").value))
        If Len(keyCell) > 0 Then
            If StrComp(keyCell, keyName, vbTextCompare) = 0 Then
                wsConfig.Cells(r, "B").value = value
                Exit Sub
            End If
        End If
    Next r
    
    ' Append new key at the next row
    r = lastRow + 1
    wsConfig.Cells(r, "A").value = keyName
    wsConfig.Cells(r, "B").value = value
End Sub

Public Sub SetupConfigValidation()
    Dim wsConfig As Worksheet
    Dim rng As Range
    
    ' Make sure Config exists and has the basic keys
    EnsureConfigSheet wsConfig
    
    '---------------------------
    ' ExportFormat (B4): plain | markdown
    '---------------------------
    Set rng = wsConfig.Range("B4")
    On Error Resume Next
    rng.Validation.Delete
    On Error GoTo 0
    With rng.Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:="plain,markdown"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ErrorTitle = "Invalid ExportFormat"
        .ErrorMessage = "Choose either 'plain' or 'markdown'."
    End With
    If Len(CStr(rng.value)) = 0 Then rng.value = "plain"
    
    '---------------------------
    ' UseArCommentPrefix (B5): TRUE | FALSE
    '---------------------------
    Set rng = wsConfig.Range("B5")
    On Error Resume Next
    rng.Validation.Delete
    On Error GoTo 0
    With rng.Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:="TRUE,FALSE"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ErrorTitle = "Invalid UseArCommentPrefix"
        .ErrorMessage = "Choose either TRUE or FALSE."
    End With
    If Len(CStr(rng.value)) = 0 Then rng.value = "FALSE"
    
    '---------------------------
    ' DefaultConfidenceLevel (B6): High | Medium | Low
    '---------------------------
    Set rng = wsConfig.Range("B6")
    On Error Resume Next
    rng.Validation.Delete
    On Error GoTo 0
    With rng.Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:="High,Medium,Low"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ErrorTitle = "Invalid DefaultConfidenceLevel"
        .ErrorMessage = "Choose High, Medium, or Low."
    End With
    If Len(CStr(rng.value)) = 0 Then rng.value = "Medium"
    
    '---------------------------
    ' UseArAuthorNames (B7): TRUE | FALSE
    '---------------------------
    Set rng = wsConfig.Range("B7")
    On Error Resume Next
    rng.Validation.Delete
    On Error GoTo 0
    With rng.Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:="TRUE,FALSE"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ErrorTitle = "Invalid UseArAuthorNames"
        .ErrorMessage = "Choose either TRUE or FALSE."
    End With
    If Len(CStr(rng.value)) = 0 Then rng.value = "FALSE"
    
    wsConfig.Columns("A:B").EntireColumn.AutoFit
    
    MsgBox "Config validation has been set up. Please review the dropdowns on the Config sheet.", _
           vbInformation, "Config Validation"
End Sub

Public Function GetConfigBool(ByVal keyName As String, _
                              Optional ByVal defaultVal As Boolean = False) As Boolean
    Dim raw As String
    raw = LCase$(Trim$(GetConfigValue(keyName, IIf(defaultVal, "TRUE", "FALSE"))))
    
    Select Case raw
        Case "true", "yes", "y", "1"
            GetConfigBool = True
        Case "false", "no", "n", "0"
            GetConfigBool = False
        Case Else
            GetConfigBool = defaultVal
    End Select
End Function



Public Sub SetupConfigAndLLMSheets()
    Dim wsConfig As Worksheet
    Dim wsPersonas As Worksheet
    
    ' Ensure Config sheet + standard keys
    EnsureConfigSheet wsConfig
    
    ' Ensure Personas sheet
    EnsurePersonasSheet wsPersonas
    
    ' Optional: set up data validation dropdowns
    SetupConfigValidation
    
    ' Ensure baseline LLM_Changes sheet exists
    SetupLLMWorkflowSheets
    
    MsgBox "Config and LLM_Changes sheets are ready.", vbInformation
End Sub

Public Sub SetupLLMWorkflowSheets()
    Dim wb As Workbook
    Dim wsChanges As Worksheet
    
    Set wb = ThisWorkbook
    
    On Error Resume Next
    Set wsChanges = wb.Worksheets("LLM_Changes")
    On Error GoTo 0
    
    If wsChanges Is Nothing Then
        Set wsChanges = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        wsChanges.name = "LLM_Changes"
        
        ' Basic headers / instructions
        With wsChanges
            .Range("A1").value = "LLM_Changes JSONL"
            .Range("A3").value = "Paste or load one JSON object per line, starting at A8."
            .Range("A5").value = "Schema (example):"
            .Range("A6").value = "{""bookmark_id"":""AR_PARA_00001"",""change_type"":""replace_text"",""new_text"":""..."",""add_comment"":""..."",""apply_change"":true,""confidence"":""Medium""}"

        End With
    End If
End Sub



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

