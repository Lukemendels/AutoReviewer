Attribute VB_Name = "modSetup"
Option Explicit

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
            .Range("A6").value = "{""bookmark_id"":""MKS_PARA_00001"",""change_type"":""replace_text"",""new_text"":""..."",""add_comment"":""..."",""apply_change"":true,""confidence"":""Medium""}"

        End With
    End If
End Sub

