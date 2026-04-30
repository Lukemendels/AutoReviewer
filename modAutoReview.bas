Attribute VB_Name = "modAutoReview"
Option Explicit

'------------------------------------------------------
' 1) Top-level macro: Start Auto-Review for a Word doc
'    - Runs ExportWordDocForLLM (new version).
'    - Reads WordDocPath & LastExportTxtPath from Config.
'    - Calls a PowerShell script to create a session folder.
'------------------------------------------------------
Public Sub AutoReview_Start()
    Dim wsConfig As Worksheet
    Dim wordPath As String
    Dim exportTxtPath As String
    Dim psScriptPath As String
    Dim cmd As String
    
    On Error GoTo ErrHandler
    
    ' Ensure Config exists
    EnsureConfigSheet wsConfig
    
    ' 1. Export Word doc for LLM (user picks the doc)
    ExportWordDocForLLM
    
    ' 2. Retrieve paths from Config (set by ExportWordDocForLLM)
    wordPath = Trim$(GetConfigValue("WordDocPath", ""))
    exportTxtPath = Trim$(GetConfigValue("LastExportTxtPath", ""))
    
    If Len(wordPath) = 0 Or Len(exportTxtPath) = 0 Then
        MsgBox "WordDocPath or LastExportTxtPath is not set in Config." & vbCrLf & _
               "Make sure ExportWordDocForLLM completed successfully.", _
               vbExclamation, "AutoReview_Start"
        Exit Sub
    End If
    
    ' 3. Absolute path to PowerShell script (adjust to your environment)
    psScriptPath = "C:\Users\Luke.Mendelsohn\OneDrive - USTSA\Documents\PowerShellScripts\Start-AutoReview.ps1"
    
    If Dir$(psScriptPath, vbNormal) = "" Then
        MsgBox "PowerShell script not found at:" & vbCrLf & psScriptPath, _
               vbExclamation, "AutoReview_Start"
        Exit Sub
    End If
    
    ' 4. Build PowerShell command line
    cmd = "powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass " & _
          "-File """ & psScriptPath & """ " & _
          "-WordDocPath """ & wordPath & """ " & _
          "-ExportTxtPath """ & exportTxtPath & """"
    
    ' DEBUG: show what we're about to run
    MsgBox "PowerShell command:" & vbCrLf & vbCrLf & cmd, vbInformation, "AutoReview_Start debug"
    
    ' 5. Launch PowerShell (non-blocking)
    Shell cmd, vbNormalFocus
    
    MsgBox "Export complete. A PowerShell window / Explorer should open " & _
           "for the LLM review session.", vbInformation, "AutoReview_Start"
    Exit Sub

ErrHandler:
    MsgBox "Error in AutoReview_Start: " & Err.Description, vbCritical
End Sub

'------------------------------------------------------
' 2) Macro: Load a JSONL file into LLM_Changes!A8:A
'    and apply edits to the Word doc
'------------------------------------------------------
Public Sub AutoReview_ApplyFromJsonFile()
    Const msoFileDialogFilePicker As Long = 3
    
    Dim wb As Workbook
    Dim wsChanges As Worksheet
    Dim fd As Object
    Dim jsonlPath As String
    
    On Error GoTo ErrHandler
    
    Set wb = ThisWorkbook
    EnsureLLMChangesSheetExists wsChanges
    
    ' 1. Ask user to pick a JSONL file
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Select JSONL file with LLM suggestions"
        .AllowMultiSelect = False
        If Len(wb.Path) > 0 Then
            .InitialFileName = wb.Path & "\"
        End If
        
        If .Show <> -1 Then
            GoTo Cleanup ' user cancelled
        End If
        
        jsonlPath = .SelectedItems(1)
    End With
    Set fd = Nothing
    
    jsonlPath = Trim$(jsonlPath)
    If Len(jsonlPath) = 0 Then GoTo Cleanup
    
    If Dir$(jsonlPath, vbNormal) = "" Then
        MsgBox "The selected JSONL file cannot be found:" & vbCrLf & jsonlPath, _
               vbCritical, "AutoReview_ApplyFromJsonFile"
        GoTo Cleanup
    End If
    
    ' 2. Load JSONL lines into LLM_Changes!A8:A
    LoadJsonlIntoLLMChanges wsChanges, jsonlPath
    
    ' 3. Apply suggestions to Word using existing engine
    ApplyWordSuggestionsFromJson
    
Cleanup:
    On Error Resume Next
    Set fd = Nothing
    Exit Sub

ErrHandler:
    MsgBox "Error in AutoReview_ApplyFromJsonFile: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

' Helper: ensure LLM_Changes exists and return it
Private Sub EnsureLLMChangesSheetExists(ByRef wsChanges As Worksheet)
    Dim wb As Workbook
    Set wb = ThisWorkbook
    
    On Error Resume Next
    Set wsChanges = wb.Worksheets("LLM_Changes")
    On Error GoTo 0
    
    If wsChanges Is Nothing Then
        SetupLLMWorkflowSheets
        Set wsChanges = wb.Worksheets("LLM_Changes")
    End If
End Sub

' Helper: read JSONL file and write each line to column A starting at row 8
Private Sub LoadJsonlIntoLLMChanges(ByVal wsChanges As Worksheet, ByVal jsonlPath As String)
    Dim stm As Object          ' ADODB.Stream
    Dim textAll As String
    Dim lines() As String
    Dim i As Long
    Dim lineText As String
    Dim row As Long
    
    On Error GoTo ErrHandler
    
    ' Clear old data starting at A8
    wsChanges.Range("A8:A" & wsChanges.Rows.Count).ClearContents
    
    ' --- Read entire file as UTF-8 text ---
    Set stm = CreateObject("ADODB.Stream")
    With stm
        .Type = 2               ' adTypeText
        .Charset = "utf-8"
        .Open
        .LoadFromFile jsonlPath
        textAll = .ReadText(-1) ' adReadAll
        .Close
    End With
    Set stm = Nothing
    
    ' Normalize line endings to vbLf
    textAll = Replace(textAll, vbCrLf, vbLf)
    textAll = Replace(textAll, vbCr, vbLf)
    
    ' Split into lines
    lines = Split(textAll, vbLf)
    
    row = 8
    For i = LBound(lines) To UBound(lines)
        lineText = lines(i)
        
        ' Strip BOM if present at start of file
        lineText = Replace(lineText, Chr$(239) & Chr$(187) & Chr$(191), "") ' EF BB BF
        lineText = Replace(lineText, ChrW$(&HFEFF), "")                     ' FE FF (Unicode BOM)
        
        lineText = Trim$(lineText)
        
        ' Skip empty or non-JSON lines (extra safety)
        If Len(lineText) > 0 Then
            If Left$(lineText, 1) = "{" Then
                wsChanges.cells(row, "A").value = lineText
                row = row + 1
            End If
        End If
    Next i
    
    MsgBox "Loaded " & (row - 8) & " JSONL lines into LLM_Changes.", _
           vbInformation, "LoadJsonlIntoLLMChanges"
    Exit Sub

ErrHandler:
    On Error Resume Next
    If Not stm Is Nothing Then
        If stm.State <> 0 Then stm.Close
    End If
    Set stm = Nothing
    MsgBox "Error loading JSONL: " & Err.Description, vbCritical, "LoadJsonlIntoLLMChanges"
End Sub




'------------------------------------------------------
' 3) Macro: Capture JSONL from clipboard via PowerShell,
'    load into LLM_Changes, and apply to Word.
'------------------------------------------------------
Public Sub AutoReview_ApplyFromClipboard()
    Dim psScriptPath As String
    Dim sessionRoot As String
    Dim lastJsonlMarker As String
    Dim jsonlPath As String
    Dim wsChanges As Worksheet
    Dim wb As Workbook
    Dim cmd As String
    Dim wsh As Object
    Dim ret As Long
    
    On Error GoTo ErrHandler
    
    Set wb = ThisWorkbook
    EnsureLLMChangesSheetExists wsChanges
    
    ' 1. Path to the PowerShell script that captures JSONL from clipboard
    psScriptPath = "C:\Users\Luke.Mendelsohn\OneDrive - USTSA\Documents\PowerShellScripts\Capture-JsonlFromClipboard.ps1" ' <-- adjust this path
    
    If Dir$(psScriptPath, vbNormal) = "" Then
        MsgBox "PowerShell script not found at:" & vbCrLf & psScriptPath, _
               vbExclamation, "AutoReview_ApplyFromClipboard"
        Exit Sub
    End If
    
    ' 2. Session root must match the one used in the PowerShell scripts
    sessionRoot = Environ$("USERPROFILE") & "\LlmSessions"
    lastJsonlMarker = sessionRoot & "\last_jsonl.txt"
    
    ' 3. Run the PowerShell script synchronously (wait for it to finish)
    cmd = "powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass " & _
          "-File """ & psScriptPath & """"
    
    Set wsh = CreateObject("WScript.Shell")
    ret = wsh.Run(cmd, 1, True) ' 1 = normal window, True = wait
    
    If ret <> 0 Then
        MsgBox "Capture-JsonlFromClipboard.ps1 returned error code " & ret & "." & vbCrLf & _
               "Make sure your clipboard contains JSONL output and try again.", _
               vbExclamation, "AutoReview_ApplyFromClipboard"
        Exit Sub
    End If
    
    ' 4. Read the JSONL path from last_jsonl.txt
    If Dir$(lastJsonlMarker, vbNormal) = "" Then
        MsgBox "Could not find last_jsonl.txt in:" & vbCrLf & sessionRoot & vbCrLf & _
               "The PowerShell script may not have completed successfully.", _
               vbExclamation, "AutoReview_ApplyFromClipboard"
        Exit Sub
    End If
    
    jsonlPath = ReadAllText(lastJsonlMarker)
    If Len(jsonlPath) = 0 Then
        MsgBox "The JSONL path recorded in last_jsonl.txt is empty or invalid.", _
               vbExclamation, "AutoReview_ApplyFromClipboard"
        Exit Sub
    End If
    
    ' DEBUG: show which JSONL file we are about to load
    MsgBox "Loading JSONL from:" & vbCrLf & jsonlPath, _
           vbInformation, "AutoReview_ApplyFromClipboard debug"

    
    ' 5. Load JSONL into LLM_Changes and apply suggestions
    LoadJsonlIntoLLMChanges wsChanges, jsonlPath
    ApplyWordSuggestionsFromJson

    Exit Sub
    
ErrHandler:
    MsgBox "Error in AutoReview_ApplyFromClipboard: " & Err.Description, vbCritical
End Sub

' Simple helper to read the entire contents of a text file
Private Function ReadAllText(ByVal fullPath As String) As String
    Dim fNum As Integer
    Dim txt As String
    
    On Error GoTo ErrHandler
    
    fNum = FreeFile
    Open fullPath For Input As #fNum
    txt = Input$(LOF(fNum), fNum)
    Close #fNum
    
    ' Remove CR/LF
    txt = Replace(txt, vbCr, "")
    txt = Replace(txt, vbLf, "")
    
    ' Strip UTF-8 BOM rendered as "﻿" (EF BB BF)
    If Left$(txt, 3) = Chr$(239) & Chr$(187) & Chr$(191) Then
        txt = Mid$(txt, 4)
    End If
    
    ' Also strip Unicode BOM &HFEFF if present
    If Left$(txt, 1) = ChrW$(&HFEFF) Then
        txt = Mid$(txt, 2)
    End If
    
    ReadAllText = Trim$(txt)
    Exit Function
    
ErrHandler:
    On Error Resume Next
    If fNum <> 0 Then Close #fNum
    ReadAllText = ""
End Function



