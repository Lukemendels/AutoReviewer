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

        wsConfig.Range("A7").value = "ActivePersona"
        wsConfig.Range("B7").value = ""

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
                val = Trim$(CStr(wsConfig.Cells(r, "B").value))
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
    
    ' Apply Modern Styling
    On Error Resume Next
    ApplyModernStyling ThisWorkbook.Worksheets("Config")
    ApplyModernStyling ThisWorkbook.Worksheets("Personas")
    ApplyModernStyling ThisWorkbook.Worksheets("LLM_Changes")
    ApplyModernStyling ThisWorkbook.Worksheets("Ratified")
    On Error GoTo 0

    MsgBox "Config, Personas, LLM_Changes, and Ratified sheets are ready and styled.", vbInformation
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

    ' Ensure baseline Ratified sheet exists (item 8: completeness gate). The
    ' operator pastes the HOT co-thinker's Turn 3 FINAL RATIFIED PACKET here,
    ' one line per row starting at A8; HandOffToSerializer reads it back to
    ' derive ExpectedEditCount/ExpectedAnchors and to embed the packet text in
    ' the serializer prompt.
    Dim wsRatified As Worksheet
    On Error Resume Next
    Set wsRatified = wb.Worksheets("Ratified")
    On Error GoTo 0

    If wsRatified Is Nothing Then
        Set wsRatified = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        wsRatified.name = "Ratified"

        With wsRatified
            .Range("A1").value = "Ratified Decision Packet"
            .Range("A3").value = "Paste the HOT co-thinker's Turn 3 FINAL RATIFIED PACKET here, one line per row, starting at A8."
            .Range("A5").value = "Each kept/fixed block must keep its original ""[n] BOOKMARK: AR_..."" line so Hand Off to Serializer can read the ratified anchors."
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
        wsPersonas.Range("B1").value = "AssistantUrl"     ' the persona's HOT co-thinker URL
        wsPersonas.Range("C1").value = "CorpusPath"
        wsPersonas.Range("D1").value = "SkillMdPath"
        wsPersonas.Range("E1").value = "TrainingDocCount"  ' redline docs mined
        wsPersonas.Range("F1").value = "LastUpdated"
        wsPersonas.Range("G1").value = "Notes"
        wsPersonas.Range("H1").value = "ExemplarCount"      ' finalized exemplars added

        ' Make headers bold
        wsPersonas.Range("A1:H1").Font.Bold = True
        wsPersonas.Columns("A:H").EntireColumn.AutoFit
    End If
End Sub

' Adds or updates a persona in the registry
Public Sub UpsertPersona(ByVal personaName As String, _
                         Optional ByVal assistantUrl As String = "", _
                         Optional ByVal corpusPath As String = "", _
                         Optional ByVal skillMdPath As String = "", _
                         Optional ByVal incrementTrainingCount As Boolean = False, _
                         Optional ByVal notes As String = "", _
                         Optional ByVal incrementExemplarCount As Boolean = False)
    Dim wsPersonas As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim foundRow As Long

    EnsurePersonasSheet wsPersonas

    ' Defensive: ensure the ExemplarCount header exists on pre-existing sheets.
    If Len(CStr(wsPersonas.Range("H1").value)) = 0 Then
        wsPersonas.Range("H1").value = "ExemplarCount"
        wsPersonas.Range("H1").Font.Bold = True
    End If

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
        wsPersonas.Cells(foundRow, "E").value = 0 ' Initialize training count
        wsPersonas.Cells(foundRow, "H").value = 0 ' Initialize exemplar count
    End If

    If Len(assistantUrl) > 0 Then wsPersonas.Cells(foundRow, "B").value = assistantUrl
    If Len(corpusPath) > 0 Then wsPersonas.Cells(foundRow, "C").value = corpusPath
    If Len(skillMdPath) > 0 Then wsPersonas.Cells(foundRow, "D").value = skillMdPath

    If incrementTrainingCount Then
        wsPersonas.Cells(foundRow, "E").value = Val(wsPersonas.Cells(foundRow, "E").value) + 1
    End If

    If incrementExemplarCount Then
        wsPersonas.Cells(foundRow, "H").value = Val(wsPersonas.Cells(foundRow, "H").value) + 1
    End If

    wsPersonas.Cells(foundRow, "F").value = Now
    If Len(notes) > 0 Then wsPersonas.Cells(foundRow, "G").value = notes

    wsPersonas.Columns("A:H").EntireColumn.AutoFit
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


Public Sub ApplyModernStyling(ByVal ws As Worksheet)
    On Error Resume Next
    
    ' Global background
    ws.Cells.Interior.Color = RGB(30, 34, 42)
    ws.Cells.Font.Name = "Segoe UI"
    ws.Cells.Font.Size = 10
    ws.Cells.Font.Color = RGB(220, 224, 230)
    
    ' Header Row Styling
    With ws.Rows(1)
        .Interior.Color = RGB(20, 24, 30)
        .Font.Bold = True
        .Font.Color = RGB(128, 190, 255)
        .RowHeight = 25
        .VerticalAlignment = xlCenter
    End With
    
    ' Draw subtle borders around the used range
    Dim lastRow As Long
    Dim lastCol As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    If lastRow > 1 And lastCol > 0 Then
        With ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Borders
            .LineStyle = xlContinuous
            .Color = RGB(60, 65, 75)
            .Weight = xlThin
        End With
    End If
    
    ' Specific sheet handling
    If ws.name = "LLM_Changes" Then
        ws.Range("A3:A6").Font.Color = RGB(160, 170, 180)
        ws.Range("A3:A6").Font.Italic = True
    ElseIf ws.name = "Ratified" Then
        ws.Range("A3:A5").Font.Color = RGB(160, 170, 180)
        ws.Range("A3:A5").Font.Italic = True
    End If

    On Error GoTo 0
End Sub

' Resolve the folder this workbook's data files (corpus, exemplars, SKILL.md,
' exports, working copies) should live in.
'
' With AutoSave on, ThisWorkbook.Path can return a SharePoint/OneDrive URL
' (https://...sharepoint.com/personal/<user>/Documents/<subpath>) instead of a
' local path. FileSystemObject hangs or raises error 52 on such URLs, and
' FileExists/FolderExists silently return False -- so every file write that
' built a path from ThisWorkbook.Path could fail. This function resolves a
' real, writable local folder and is the single source of truth for "where do
' this workbook's files live".
Public Function GetWorkFolder() As String
    Dim fso As Object
    Dim cfg As String
    Dim p As String
    Dim idx As Long
    Dim remainder As String
    Dim oneDrive As String
    Dim candidate As String

    Set fso = CreateObject("Scripting.FileSystemObject")

    ' 1) An explicit, verified override always wins.
    cfg = Trim$(GetConfigValue("WorkFolder", ""))
    If Len(cfg) > 0 Then
        If fso.FolderExists(cfg) Then
            GetWorkFolder = cfg
            Exit Function
        End If
    End If

    p = ThisWorkbook.Path

    ' 2) A normal local (or UNC) path: use it as-is.
    If LCase$(Left$(p, 4)) <> "http" Then
        GetWorkFolder = p
        Exit Function
    End If

    ' 3) URL translation: .../personal/<user>/Documents/<subpath> ->
    '    %OneDriveCommercial%\Documents\<subpath> (or %OneDrive% as a fallback).
    idx = InStr(1, p, "/Documents", vbTextCompare)
    If idx > 0 Then
        remainder = Mid$(p, idx + Len("/Documents"))
        remainder = Replace(remainder, "%20", " ")
        remainder = Replace(remainder, "/", "\")

        oneDrive = Environ$("OneDriveCommercial")
        If Len(oneDrive) = 0 Then oneDrive = Environ$("OneDrive")

        If Len(oneDrive) > 0 Then
            candidate = oneDrive & "\Documents" & remainder
            ' 4) Verify and persist (visible/auditable on the Config sheet).
            If fso.FolderExists(candidate) Then
                SetConfigValue "WorkFolder", candidate
                GetWorkFolder = candidate
                Exit Function
            End If
        End If
    End If

    ' 5) Fall back to a folder under the user's profile, creating it if needed.
    candidate = Environ$("USERPROFILE") & "\AutoReviewer"
    If Not fso.FolderExists(candidate) Then
        fso.CreateFolder candidate
        MsgBox "Could not resolve a local OneDrive folder for this workbook's " & _
               "SharePoint/OneDrive location." & vbCrLf & vbCrLf & _
               "AutoReviewer files (corpus, exemplars, exports, etc.) will be " & _
               "saved to:" & vbCrLf & candidate & vbCrLf & vbCrLf & _
               "You can change this later via ""Set Work Folder"" on the dashboard.", _
               vbInformation, "AutoReviewer Work Folder"
    End If
    SetConfigValue "WorkFolder", candidate
    GetWorkFolder = candidate
End Function

' Sets and reads back 3 Config keys (including "ActivePersona") to verify
' GetConfigValue/SetConfigValue round-trip correctly: an exact, case-insensitive,
' Trim$ key match against column A, with the value read from the adjacent column
' B -- never a positional cell, and never a value that "happens" to belong to a
' different key.
Public Sub TestConfigRoundtrip()
    Dim results As String
    Dim allPass As Boolean
    allPass = True

    allPass = CheckConfigRoundtripKey("ActivePersona", "AR_TestPersona_" & Format$(Now, "hhnnss"), results) And allPass
    allPass = CheckConfigRoundtripKey("AR_TestKeyAlpha", "alpha-value-1", results) And allPass
    allPass = CheckConfigRoundtripKey("AR_TestKeyBeta", "beta value, with spaces", results) And allPass

    If allPass Then
        MsgBox "TestConfigRoundtrip: PASS" & vbCrLf & vbCrLf & results, vbInformation, "Config Roundtrip"
    Else
        MsgBox "TestConfigRoundtrip: FAIL" & vbCrLf & vbCrLf & results, vbCritical, "Config Roundtrip"
    End If
End Sub

' Helper for TestConfigRoundtrip: sets keyName=value, reads it back, and checks
' that an UNRELATED key (a different case/whitespace variant of the same name)
' still resolves to the SAME value -- proving the match is on the trimmed,
' case-insensitive key, not a positional/coincidental cell.
Private Function CheckConfigRoundtripKey(ByVal keyName As String, ByVal value As String, _
                                          ByRef results As String) As Boolean
    Dim readBack As String
    Dim readBackVariant As String
    Dim ok As Boolean

    SetConfigValue keyName, value
    readBack = GetConfigValue(keyName, "<MISSING>")
    readBackVariant = GetConfigValue("  " & UCase$(keyName) & "  ", "<MISSING>")

    ok = (readBack = value) And (readBackVariant = value)

    results = results & keyName & ": set=""" & value & """ got=""" & readBack & _
              """ (case/whitespace variant got=""" & readBackVariant & """) -> " & _
              IIf(ok, "OK", "FAIL") & vbCrLf

    CheckConfigRoundtripKey = ok
End Function
