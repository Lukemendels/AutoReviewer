Attribute VB_Name = "ConfigHelpers"
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
        
        wsConfig.Range("A5").value = "UseMksCommentPrefix"
        wsConfig.Range("B5").value = "FALSE"     ' TRUE | FALSE
        
        wsConfig.Range("A6").value = "DefaultConfidenceLevel"
        wsConfig.Range("B6").value = "Medium"    ' High | Medium | Low
        
        wsConfig.Range("A7").value = "UseMksAuthorNames"
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
    ' UseMksCommentPrefix (B5): TRUE | FALSE
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
        .ErrorTitle = "Invalid UseMksCommentPrefix"
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
    ' UseMksAuthorNames (B7): TRUE | FALSE
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
        .ErrorTitle = "Invalid UseMksAuthorNames"
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

