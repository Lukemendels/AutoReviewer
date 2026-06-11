Attribute VB_Name = "modDashboardUI"
Option Explicit

' Layout grid for the dashboard. Every button is placed via PlaceButton(col,
' row, ...), which derives its Top/Left from these constants -- no hardcoded
' coordinates anywhere else, so columns of any length can never collide.
Private Const GRID_LEFT As Double = 30
Private Const GRID_TOP As Double = 130
Private Const COL_WIDTH As Double = 280
Private Const COL_GAP As Double = 30
Private Const BTN_HEIGHT As Double = 45
Private Const ROW_GAP As Double = 12

Public Sub BuildDashboard()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim shp As Shape

    ' First, ensure sheets are set up
    modAppCore.SetupConfigAndLLMSheets

    Set wb = ThisWorkbook

    ' 1. Delete existing Dashboard if it exists. This also discards every shape
    ' on it, so re-running BuildDashboard can never leave orphaned/stacked
    ' buttons from a prior layout.
    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Worksheets("Dashboard").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    ' 2. Create new Dashboard sheet
    Set ws = wb.Worksheets.Add(Before:=wb.Worksheets(1))
    ws.name = "Dashboard"

    ' 3. Apply Base Styling (Dark Mode)
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False

    ' Dark slate background
    ws.Cells.Interior.Color = RGB(20, 24, 30)

    ws.Columns("A:A").ColumnWidth = 5
    ws.Columns("B:B").ColumnWidth = 50
    ws.Columns("C:C").ColumnWidth = 50

    ' 4. Add Title
    Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, 30, 20, 400, 60)
    With shp
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        With .TextFrame2.TextRange
            .Text = "AutoReviewer Dashboard V2"
            .Font.Name = "Segoe UI"
            .Font.Size = 28
            .Font.Bold = msoTrue
            .Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        End With
    End With

    ' 5. Active Persona Status
    ws.Cells(5, 2).value = "Active Persona: " & modAppCore.GetConfigValue("ActivePersona", "None")
    ws.Cells(5, 2).Font.Name = "Segoe UI"
    ws.Cells(5, 2).Font.Size = 14
    ws.Cells(5, 2).Font.Bold = True
    ws.Cells(5, 2).Font.Color = RGB(160, 170, 180)

    ' --- COLUMN 1: TRAIN NEW PERSONA ---
    PlaceColumnHeader ws, 1, "TRAIN PERSONA", RGB(128, 90, 213) ' Purple

    PlaceButton ws, 1, 1, "1. Set Active Persona", "modDashboardUI.SetActivePersona", RGB(74, 85, 104)
    PlaceButton ws, 1, 2, "2a. Add Doc to Corpus (redlines)", "modTrainingPipeline.AddDocToCorpus", RGB(128, 90, 213)
    PlaceButton ws, 1, 3, "2b. Add Finalized Exemplar", "modTrainingPipeline.AddFinalizedExemplar", RGB(128, 90, 213)
    PlaceButton ws, 1, 4, "3. Reduce Pass 1: Cluster", "modTrainingPipeline.RunReducePass1", RGB(128, 90, 213)
    PlaceButton ws, 1, 5, "4. Reduce Pass 2: Heuristics", "modTrainingPipeline.RunReducePass2", RGB(128, 90, 213)
    PlaceButton ws, 1, 6, "5. Reduce Pass 3: SKILL.md", "modTrainingPipeline.RunReducePass3", RGB(128, 90, 213)
    PlaceButton ws, 1, 7, "6. Save SKILL.md", "modTrainingPipeline.SaveSkillMd", RGB(128, 90, 213)

    ' --- COLUMN 2: RUN REVIEW ---
    PlaceColumnHeader ws, 2, "RUN REVIEW", RGB(49, 130, 206) ' Blue

    PlaceButton ws, 2, 1, "1. Select Persona for Review", "modDashboardUI.SetActivePersona", RGB(74, 85, 104)
    PlaceButton ws, 2, 2, "2. Prepare for Review (Co-thinker)", "modReviewExport.ExportWordDocForLLM", RGB(49, 130, 206)
    PlaceButton ws, 2, 3, "3. Hand off to Serializer", "modReviewExport.HandOffToSerializer", RGB(214, 158, 46) ' Amber: ratify between 2 and 3
    PlaceButton ws, 2, 4, "4. Apply LLM Edits to Word", "modReviewImport.ApplyWordSuggestionsFromJson", RGB(56, 161, 105) ' Green

    ' --- COLUMN 3: RESPOND TO REVIEW ---
    PlaceColumnHeader ws, 3, "RESPOND TO REVIEW", RGB(221, 107, 32) ' Orange

    ' Incorporation uses the shared Incorporator assistant -- no persona needed.
    PlaceButton ws, 3, 1, "1. Export Document for Feedback", "modReviewExport.ExportWordDocForRespondMode", RGB(221, 107, 32)
    PlaceButton ws, 3, 2, "2. Hand off to Serializer", "modReviewExport.HandOffToSerializer", RGB(214, 158, 46) ' Amber: ratify between 1 and 2
    PlaceButton ws, 3, 3, "3. Apply LLM Edits to Word", "modReviewImport.ApplyWordSuggestionsFromJson", RGB(56, 161, 105) ' Green

    ' --- COLUMN 4: CONFIG / UTILITIES ---
    ' All settings/config buttons live in their own column so they can never
    ' overlap a workflow column regardless of how long that column grows.
    PlaceColumnHeader ws, 4, "CONFIG / UTILITIES", RGB(160, 170, 180) ' Gray

    PlaceButton ws, 4, 1, "Go to Config", "modDashboardUI.GoToConfigSheet", RGB(74, 85, 104)
    PlaceButton ws, 4, 2, "Set Work Folder", "modDashboardUI.SetWorkFolder", RGB(74, 85, 104)
    PlaceButton ws, 4, 3, "Set Chat URL", "modDashboardUI.SetChatUrl", RGB(74, 85, 104)
    PlaceButton ws, 4, 4, "Set Serializer URL", "modDashboardUI.SetSerializerUrl", RGB(74, 85, 104)
    PlaceButton ws, 4, 5, "Set Incorporator URL", "modDashboardUI.SetIncorporatorUrl", RGB(74, 85, 104)
    PlaceButton ws, 4, 6, "Set Researcher URL", "modDashboardUI.SetResearcherUrl", RGB(74, 85, 104)
    PlaceButton ws, 4, 7, "Set Citation URL", "modDashboardUI.SetCitationUrl", RGB(74, 85, 104)
    PlaceButton ws, 4, 8, "Open Researcher", "modDashboardUI.OpenResearcher", RGB(74, 85, 104)

    ws.Range("A1").Select
End Sub

' Place button `row` (1-based, top-down) in grid column `col` (1-based,
' left-to-right). Top/Left are derived purely from col/row plus the grid
' constants, so two calls with different (col,row) can never produce
' overlapping shapes.
Private Sub PlaceButton(ByVal ws As Worksheet, ByVal col As Long, ByVal row As Long, _
                        ByVal caption As String, ByVal macroName As String, ByVal bgColor As Long)
    Dim leftPos As Double
    Dim topPos As Double

    leftPos = GRID_LEFT + (col - 1) * (COL_WIDTH + COL_GAP)
    topPos = GRID_TOP + (row - 1) * (BTN_HEIGHT + ROW_GAP)

    CreateModernButton ws, leftPos, topPos, COL_WIDTH, BTN_HEIGHT, caption, macroName, bgColor
End Sub

' Place the section-title textbox above row 1 of grid column `col`.
Private Sub PlaceColumnHeader(ByVal ws As Worksheet, ByVal col As Long, ByVal caption As String, ByVal textColor As Long)
    Dim shp As Shape
    Dim leftPos As Double

    leftPos = GRID_LEFT + (col - 1) * (COL_WIDTH + COL_GAP)

    Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, leftPos, GRID_TOP - 30, COL_WIDTH + 20, 30)
    shp.Fill.Visible = msoFalse
    shp.Line.Visible = msoFalse
    shp.TextFrame2.TextRange.Text = caption
    shp.TextFrame2.TextRange.Font.Name = "Segoe UI"
    shp.TextFrame2.TextRange.Font.Size = 14
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = textColor
End Sub

Public Sub SetResearcherUrl()
    Dim current As String
    Dim choice As String

    current = modAppCore.GetConfigValue("ResearcherUrl", "")
    choice = InputBox("Enter the shared Researcher assistant URL." & vbCrLf & vbCrLf & _
                      "One generic assistant (set up once from " & _
                      "TEMPLATE_SKILL_RESEARCHER.md, with the citation standard pasted in) " & _
                      "that runs focused data/citation side-investigations from a research " & _
                      "brief. Shared across all documents; not a persona.", _
                      "Set Researcher URL", current)

    If Trim$(choice) <> "" Then
        modAppCore.SetConfigValue "ResearcherUrl", Trim$(choice)
        MsgBox "Researcher URL saved.", vbInformation
    End If
End Sub

Public Sub SetCitationUrl()
    Dim current As String
    Dim choice As String

    current = modAppCore.GetConfigValue("CitationUrl", "")
    choice = InputBox("Enter the shared Citation assistant URL." & vbCrLf & vbCrLf & _
                      "One generic assistant (set up once from " & _
                      "TEMPLATE_SKILL_CITATION.md) for 'I just need a citation' moments. " & _
                      "Same canonical standard the Researcher embeds.", _
                      "Set Citation URL", current)

    If Trim$(choice) <> "" Then
        modAppCore.SetConfigValue "CitationUrl", Trim$(choice)
        MsgBox "Citation URL saved.", vbInformation
    End If
End Sub

Public Sub OpenResearcher()
    Dim url As String
    url = Trim$(modAppCore.GetConfigValue("ResearcherUrl", ""))
    If Len(url) = 0 Then
        MsgBox "No Researcher URL is set. Use 'Set Researcher URL' first.", vbExclamation
        Exit Sub
    End If
    modSysUtils.OpenURL url
End Sub

Public Sub SetIncorporatorUrl()
    Dim current As String
    Dim choice As String

    current = modAppCore.GetConfigValue("IncorporatorUrl", "")
    choice = InputBox("Enter the shared Incorporator assistant URL." & vbCrLf & vbCrLf & _
                      "This is one generic assistant (set up once from " & _
                      "TEMPLATE_SKILL_INCORPORATOR.md) that helps you understand and act " & _
                      "on reviewer feedback. It is shared across all documents and is not " & _
                      "tied to any persona.", _
                      "Set Incorporator URL", current)

    If Trim$(choice) <> "" Then
        modAppCore.SetConfigValue "IncorporatorUrl", Trim$(choice)
        MsgBox "Incorporator URL saved.", vbInformation
    End If
End Sub

Public Sub SetSerializerUrl()
    Dim current As String
    Dim choice As String

    current = modAppCore.GetConfigValue("SerializerUrl", "")
    choice = InputBox("Enter the shared COLD Serializer assistant URL." & vbCrLf & vbCrLf & _
                      "This is a single, generic assistant that holds the JSONL output " & _
                      "contract and is shared across all personas. The per-persona HOT " & _
                      "co-thinker URL lives in the Personas sheet (AssistantUrl column).", _
                      "Set Serializer URL", current)

    If Trim$(choice) <> "" Then
        modAppCore.SetConfigValue "SerializerUrl", Trim$(choice)
        MsgBox "Serializer URL saved.", vbInformation
    End If
End Sub

' "Set Chat URL": the URL Reduce Pass 1 opens (a fresh, no-assistant DHSChat
' chat used to cluster the persona's training corpus). Stored as the
' "CustomGptUrl" Config key, which RunReducePass1 already reads.
Public Sub SetChatUrl()
    Dim current As String
    Dim choice As String

    current = modAppCore.GetConfigValue("CustomGptUrl", "")
    choice = InputBox("Enter the Chat URL for Reduce Pass 1." & vbCrLf & vbCrLf & _
                      "This opens a FRESH chat (no assistant) where you drag in the " & _
                      "training corpus/exemplars and run the Reduce passes. Shared " & _
                      "across all personas; stored as the 'CustomGptUrl' Config key.", _
                      "Set Chat URL", current)

    If Trim$(choice) <> "" Then
        modAppCore.SetConfigValue "CustomGptUrl", Trim$(choice)
        MsgBox "Chat URL saved.", vbInformation
    End If
End Sub

' "Set Work Folder": override where corpus files, exemplars, SKILL.md, and
' exports are written (modAppCore.GetWorkFolder). Useful when the workbook
' lives on a SharePoint/OneDrive URL path and the auto-detected local sync
' folder is not the desired location.
Public Sub SetWorkFolder()
    Dim fso As Object
    Dim current As String
    Dim choice As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    current = modAppCore.GetWorkFolder()

    choice = InputBox("Enter the local folder where AutoReviewer should store " & _
                      "corpus files, exemplars, SKILL.md, and exports." & vbCrLf & vbCrLf & _
                      "The folder must already exist. Leave unchanged to keep the " & _
                      "current folder.", _
                      "Set Work Folder", current)

    If Trim$(choice) = "" Then Exit Sub
    choice = Trim$(choice)

    If Not fso.FolderExists(choice) Then
        MsgBox "That folder does not exist:" & vbCrLf & choice, vbExclamation, "Set Work Folder"
        Exit Sub
    End If

    modAppCore.SetConfigValue "WorkFolder", choice
    MsgBox "Work folder saved:" & vbCrLf & choice, vbInformation, "Set Work Folder"
End Sub

Public Sub SetActivePersona()
    Dim current As String
    Dim choice As String

    current = modAppCore.GetConfigValue("ActivePersona", "")
    choice = InputBox("Enter the Persona Name:", "Set Active Persona", current)

    If Trim(choice) <> "" Then
        modAppCore.SetConfigValue "ActivePersona", choice
        modAppCore.UpsertPersona choice
        ' Rebuild dashboard to update label
        BuildDashboard
    End If
End Sub

Private Sub CreateModernButton(ws As Worksheet, Left As Double, Top As Double, Width As Double, Height As Double, _
                               Text As String, MacroName As String, BgColor As Long)
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, Left, Top, Width, Height)
    With shp
        .Fill.ForeColor.RGB = BgColor
        .Line.Visible = msoFalse
        .Adjustments.Item(1) = 0.15
        .Shadow.Type = msoShadow21
        .Shadow.ForeColor.RGB = RGB(0, 0, 0)
        .Shadow.Transparency = 0.6
        .Shadow.Size = 100
        .Shadow.Blur = 8
        .Shadow.OffsetX = 0
        .Shadow.OffsetY = 4
        With .TextFrame2
            .VerticalAnchor = msoAnchorMiddle
            With .TextRange
                .Text = Text
                .Font.Name = "Segoe UI SemiBold"
                .Font.Size = 12
                .Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
                .ParagraphFormat.Alignment = msoAlignCenter
            End With
        End With
        .OnAction = MacroName
    End With
End Sub

Public Sub GoToConfigSheet()
    Dim wsConfig As Worksheet
    On Error Resume Next
    Set wsConfig = ThisWorkbook.Worksheets("Config")
    On Error GoTo 0
    If Not wsConfig Is Nothing Then
        wsConfig.Activate
    Else
        MsgBox "Config sheet not found. Please run a command that initializes it first.", vbExclamation
    End If
End Sub
