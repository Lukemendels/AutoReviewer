Attribute VB_Name = "ExportToJSON"
Option Explicit

'====================================================================
'  Global-scope constants – change here only
'====================================================================
Const MAX_FILES            As Long = 300      'Failsafe cap on number of output files
Const TOKEN_RATIO          As Long = 4        '˜ 4 characters per token
Const TOKENS_WORKSPACES    As Long = 700      '˜ 2,800 characters
Const TOKENS_STD_DHSCHAT   As Long = 120000   '˜ 480,000 characters
Const SRC_SHEET_NAME       As String = "Data" 'Worksheet holding the table
Const SRC_TABLE_NAME       As String = "Table1"

'====================================================================
'  Main entry point
'====================================================================
Sub RunExport()

    '--------------- procedure-level variables ----------------------
    Dim ws          As Object      'Worksheet (late-bound)
    Dim lo          As Object      'ListObject (late-bound)
    Dim rng         As Object      'Range (late-bound)
    Dim fmtPretty   As Boolean     'True = JSON array; False = JSONL
    Dim useChunk    As Boolean
    Dim chunkTokens As Long
    Dim chunkChars  As Long
    Dim totalChars  As Long
    Dim estFiles    As Long
    Dim folderPath  As String
    Dim baseName    As String
    Dim resp        As String
    Dim ans         As VbMsgBoxResult
    
    '----------------------------------------------------------------
    ' 0. Locate or create the source table
    '----------------------------------------------------------------
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SRC_SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Worksheet '" & SRC_SHEET_NAME & "' was not found.", vbExclamation
        Exit Sub
    End If
    
    'Remove prior table and clear formats
    On Error Resume Next
    ws.ListObjects(SRC_TABLE_NAME).Unlist
    On Error GoTo 0
    ws.UsedRange.ClearFormats
    
    If Application.WorksheetFunction.CountA(ws.Cells) = 0 Then
        MsgBox "'" & SRC_SHEET_NAME & "' sheet is empty—nothing to export.", vbExclamation
        Exit Sub
    End If
    
    'Create a fresh table covering the used range
    Dim lastRow As Long, lastCol As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(-4162).row         'xlUp = -4162
    lastCol = ws.Cells(1, ws.Columns.Count).End(-4159).Column    'xlToLeft = -4159
    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
    
    'ListObjects.Add(SourceType:=xlSrcRange, Source:=rng, XlListObjectHasHeaders:=xlYes)
    Set lo = ws.ListObjects.Add(1, rng, , 1) 'xlSrcRange=1, xlYes=1
    lo.name = SRC_TABLE_NAME
    
    '----------------------------------------------------------------
    ' 1. Formatting choice: Pretty vs. Compact (Compact = JSONL)
    '----------------------------------------------------------------
    ans = MsgBox("Pretty-printed JSON (with line breaks and indentation)?" & _
                 vbCrLf & "Yes = Pretty (JSON array)    No = Compact (JSONL)", _
                 vbYesNo + vbQuestion, "Formatting")
    fmtPretty = (ans = vbYes)
    
    '----------------------------------------------------------------
    ' 2. Chunking options (UI speaks in TOKENS)
    '----------------------------------------------------------------
    useChunk = (MsgBox("Split output into multiple files?", _
                       vbYesNo + vbQuestion, "Chunking") = vbYes)
    
    If useChunk Then
        resp = InputBox( _
            "Select chunk size (enter 1, 2, or 3):" & vbCrLf & vbCrLf & _
            "1  –  " & Format$(TOKENS_WORKSPACES, "#,##0") & " tokens  (Workspaces)" & vbCrLf & _
            "2  –  " & Format$(TOKENS_STD_DHSCHAT, "#,##0") & " tokens  (Std. DHSChat max)" & vbCrLf & _
            "3  –  Custom token count", _
            "Chunk Size (in tokens)", "1")
        
        If resp = vbNullString Then Exit Sub
        
        Select Case Trim$(resp)
            Case "1": chunkTokens = TOKENS_WORKSPACES
            Case "2": chunkTokens = TOKENS_STD_DHSCHAT
            Case "3"
                chunkTokens = CLng(Application.InputBox( _
                               Prompt:="Enter maximum tokens per file:", _
                               Title:="Custom Token Size", Type:=1))
                If chunkTokens <= 0 Then
                    MsgBox "Invalid number. Defaulting to " & _
                           Format$(TOKENS_STD_DHSCHAT, "#,##0") & " tokens.", vbInformation
                    chunkTokens = TOKENS_STD_DHSCHAT
                End If
            Case Else
                MsgBox "Invalid selection. Export cancelled.", vbExclamation
                Exit Sub
        End Select
        
        chunkChars = chunkTokens * TOKEN_RATIO  'convert tokens to characters estimate
        
        'Estimate file count in the chosen format
        totalChars = Len(JSONFromTable(lo, fmtPretty))
        estFiles = (totalChars + chunkChars - 1) \ chunkChars
        
        If estFiles > MAX_FILES Then
            MsgBox "That chunk size would create about " & _
                   Format$(estFiles, "#,##0") & " files (limit " & MAX_FILES & ")." & vbCrLf & _
                   "Please choose a larger chunk size.", vbExclamation
            Exit Sub
        End If
    End If
    
    '----------------------------------------------------------------
    ' 3. Destination folder
    '----------------------------------------------------------------
    With Application.FileDialog(4) 'msoFileDialogFolderPicker = 4
        .Title = "Select destination folder"
        If .Show <> -1 Then Exit Sub
        folderPath = .SelectedItems(1)
        If Right$(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    End With
    
    '----------------------------------------------------------------
    ' 4. Base file name
    '----------------------------------------------------------------
    baseName = InputBox("Enter base file name (no extension):", _
                        "File Name", "TableExport")
    If baseName = vbNullString Then Exit Sub
    baseName = Replace(Trim$(baseName), ".", "")
    If baseName = "" Then baseName = "TableExport"
    
    '----------------------------------------------------------------
    ' 5. Export (JSON array if pretty; JSONL if compact)
    '----------------------------------------------------------------
    If useChunk Then
        ExportWithChunking lo, chunkChars, folderPath, baseName, fmtPretty
    Else
        If fmtPretty Then
            WriteTextToFileUTF8NoBOM folderPath & baseName & ".json", _
                                      JSONFromTable(lo, True)
        Else
            WriteTextToFileUTF8NoBOM folderPath & baseName & ".jsonl", _
                                      JSONFromTable(lo, False)
        End If
    End If
    
    '----------------------------------------------------------------
    ' 6. Reveal destination folder
    '----------------------------------------------------------------
    CreateObject("WScript.Shell").Run _
        "explorer.exe """ & folderPath & """", 1, False
End Sub

'====================================================================
'  Build JSON string from the table
'    - pretty=True  -> JSON array (pretty-printed)
'    - pretty=False -> JSONL (one compact JSON object per line)
'====================================================================
Private Function JSONFromTable(lo As Object, ByVal pretty As Boolean) As String
    Dim hasRows As Boolean
    hasRows = Not (lo.DataBodyRange Is Nothing)
    
    Dim headers() As String
    Dim c As Long
    ReDim headers(1 To lo.ListColumns.Count)
    For c = 1 To lo.ListColumns.Count
        headers(c) = EscapeJSON(lo.ListColumns(c).name)
    Next c
    
    If Not hasRows Then
        If pretty Then
            JSONFromTable = "[]"
        Else
            JSONFromTable = "" 'empty JSONL when no rows
        End If
        Exit Function
    End If
    
    Dim data As Variant
    data = lo.DataBodyRange.value
    
    Dim r As Long, out As String
    If pretty Then
        Dim indent1 As String, indent2 As String, nl As String
        indent1 = "  ": indent2 = "    ": nl = vbCrLf
        
        out = "[" & nl
        For r = 1 To UBound(data, 1)
            out = out & indent1 & "{" & nl
            For c = 1 To UBound(data, 2)
                out = out & indent2 & """" & headers(c) & """: """ & _
                      EscapeJSON(data(r, c)) & """"
                If c < UBound(data, 2) Then out = out & "," & nl
            Next c
            out = out & nl & indent1 & "}"
            If r < UBound(data, 1) Then out = out & "," & nl
        Next r
        out = out & nl & "]"
        JSONFromTable = out
    Else
        'Compact JSONL: one object per line, no surrounding array
        Dim lines() As String
        ReDim lines(1 To UBound(data, 1))
        Dim obj As String
        For r = 1 To UBound(data, 1)
            obj = "{"
            For c = 1 To UBound(data, 2)
                obj = obj & """" & headers(c) & """:""" & EscapeJSON(data(r, c)) & """"
                If c < UBound(data, 2) Then obj = obj & ","
            Next c
            obj = obj & "}"
            lines(r) = obj
        Next r
        JSONFromTable = Join(lines, vbCrLf)
    End If
End Function

'====================================================================
'  Write multiple chunks:
'    - pretty=True  -> JSON array split across .json files
'    - pretty=False -> JSONL split across .jsonl files
'====================================================================
Private Sub ExportWithChunking(lo As Object, ByVal MaxLen As Long, _
                               ByVal folderPath As String, ByVal baseName As String, _
                               ByVal pretty As Boolean)
    
    Dim hasRows As Boolean
    hasRows = Not (lo.DataBodyRange Is Nothing)
    
    Dim headers() As String
    Dim c As Long
    ReDim headers(1 To lo.ListColumns.Count)
    For c = 1 To lo.ListColumns.Count
        headers(c) = EscapeJSON(lo.ListColumns(c).name)
    Next c
    
    Dim fileIdx As Long
    fileIdx = 1
    
    If Not hasRows Then
        'No data rows; emit empty container
        If pretty Then
            WriteTextToFileUTF8NoBOM folderPath & baseName & "_" & fileIdx & ".json", "[]"
        Else
            WriteTextToFileUTF8NoBOM folderPath & baseName & "_" & fileIdx & ".jsonl", ""
        End If
        Exit Sub
    End If
    
    Dim data As Variant
    data = lo.DataBodyRange.value
    
    Dim r As Long, obj As String, buffer As String
    
    If pretty Then
        '----- JSON array chunking -----
        Dim indent1 As String, indent2 As String, nl As String
        indent1 = "  ": indent2 = "    ": nl = vbCrLf
        
        buffer = "[" & nl
        For r = 1 To UBound(data, 1)
            obj = indent1 & "{" & nl
            For c = 1 To UBound(data, 2)
                obj = obj & indent2 & """" & headers(c) & """: """ & _
                      EscapeJSON(data(r, c)) & """"
                If c < UBound(data, 2) Then obj = obj & "," & nl
            Next c
            obj = obj & nl & indent1 & "}"
            If r < UBound(data, 1) Then obj = obj & "," & nl
            
            If Len(buffer) + Len(obj) + Len(nl) + 1 > MaxLen Then   ' +1 for closing ]
                buffer = buffer & nl & "]"
                WriteTextToFileUTF8NoBOM folderPath & baseName & "_" & fileIdx & ".json", buffer
                fileIdx = fileIdx + 1
                buffer = "[" & nl & obj
            Else
                buffer = buffer & obj
            End If
        Next r
        
        buffer = buffer & nl & "]"
        WriteTextToFileUTF8NoBOM folderPath & baseName & "_" & fileIdx & ".json", buffer
    Else
        '----- JSONL chunking -----
        buffer = ""
        For r = 1 To UBound(data, 1)
            obj = "{"
            For c = 1 To UBound(data, 2)
                obj = obj & """" & headers(c) & """:""" & EscapeJSON(data(r, c)) & """"
                If c < UBound(data, 2) Then obj = obj & ","
            Next c
            obj = obj & "}"
            
            If Len(buffer) = 0 Then
                If Len(obj) > MaxLen Then
                    WriteTextToFileUTF8NoBOM folderPath & baseName & "_" & fileIdx & ".jsonl", obj
                    fileIdx = fileIdx + 1
                Else
                    buffer = obj
                End If
            Else
                If Len(buffer) + Len(vbCrLf) + Len(obj) > MaxLen Then
                    WriteTextToFileUTF8NoBOM folderPath & baseName & "_" & fileIdx & ".jsonl", buffer
                    fileIdx = fileIdx + 1
                    If Len(obj) > MaxLen Then
                        WriteTextToFileUTF8NoBOM folderPath & baseName & "_" & fileIdx & ".jsonl", obj
                        fileIdx = fileIdx + 1
                        buffer = ""
                    Else
                        buffer = obj
                    End If
                Else
                    buffer = buffer & vbCrLf & obj
                End If
            End If
        Next r
        
        If Len(buffer) > 0 Then
            WriteTextToFileUTF8NoBOM folderPath & baseName & "_" & fileIdx & ".jsonl", buffer
        End If
    End If
End Sub

'====================================================================
'  Write text to disk as UTF-8 without BOM (late-bound ADODB.Stream)
'====================================================================
Private Sub WriteTextToFileUTF8NoBOM(ByVal fullPath As String, ByVal txt As String)
    Dim stmText As Object   ' ADODB.Stream (text)
    Dim stmBin  As Object   ' ADODB.Stream (binary)
    
    Set stmText = CreateObject("ADODB.Stream")
    With stmText
        .Type = 2                   'adTypeText
        .Charset = "utf-8"
        .Open
        .WriteText txt
        .Position = 0
        .Type = 1                   'adTypeBinary
        If .Size >= 3 Then .Position = 3  'skip UTF-8 BOM (EF BB BF)
    End With
    
    Set stmBin = CreateObject("ADODB.Stream")
    With stmBin
        .Type = 1                   'adTypeBinary
        .Open
        stmText.CopyTo stmBin       'copy bytes sans BOM
        .SaveToFile fullPath, 2     'adSaveCreateOverWrite
        .Close
    End With
    
    stmText.Close
End Sub

'==================================================================
' Escape backslashes, quotes, and remove embedded line breaks
'==================================================================
Private Function EscapeJSON(v As Variant) As String
    Dim s As String
    On Error Resume Next
    s = CStr(v)
    On Error GoTo 0
    
    s = Replace(s, vbCrLf, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbTab, " ")
    
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    
    EscapeJSON = s
End Function

