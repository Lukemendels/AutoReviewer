Attribute VB_Name = "modDashboardUI"
Option Explicit

Public Sub BuildDashboard()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim shp As Shape
    Dim btnTop As Double
    Dim btnLeft As Double
    Dim btnWidth As Double
    Dim btnHeight As Double
    
    Set wb = ThisWorkbook
    
    ' 1. Delete existing Dashboard if it exists
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
    ws.Cells(5, 2).value = "Active Persona: " & ConfigHelpers.GetConfigValue("ActivePersona", "None")
    ws.Cells(5, 2).Font.Name = "Segoe UI"
    ws.Cells(5, 2).Font.Size = 14
    ws.Cells(5, 2).Font.Bold = True
    ws.Cells(5, 2).Font.Color = RGB(160, 170, 180)
    
    ' Button Layout Params
    btnWidth = 280
    btnHeight = 45
    
    ' --- GROUP 1: TRAIN NEW PERSONA ---
    btnLeft = 30
    btnTop = 130
    
    Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, btnLeft, btnTop - 30, 300, 30)
    shp.Fill.Visible = msoFalse
    shp.Line.Visible = msoFalse
    shp.TextFrame2.TextRange.Text = "TRAIN PERSONA"
    shp.TextFrame2.TextRange.Font.Name = "Segoe UI"
    shp.TextFrame2.TextRange.Font.Size = 14
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(128, 90, 213) ' Purple
    
    CreateModernButton ws, btnLeft, btnTop, btnWidth, btnHeight, _
        "1. Set Active Persona", "modDashboardUI.SetActivePersona", RGB(74, 85, 104)
    btnTop = btnTop + btnHeight + 15
        
    CreateModernButton ws, btnLeft, btnTop, btnWidth, btnHeight, _
        "2. Add Doc to Corpus", "modTrainingCorpusBuilder.AddDocToCorpus", RGB(128, 90, 213)
    btnTop = btnTop + btnHeight + 15
        
    CreateModernButton ws, btnLeft, btnTop, btnWidth, btnHeight, _
        "3. Reduce Pass 1: Cluster", "modTrainingOrchestrator.RunReducePass1", RGB(128, 90, 213)
    btnTop = btnTop + btnHeight + 10
        
    CreateModernButton ws, btnLeft, btnTop, btnWidth, btnHeight, _
        "4. Reduce Pass 2: Heuristics", "modTrainingOrchestrator.RunReducePass2", RGB(128, 90, 213)
    btnTop = btnTop + btnHeight + 10
    
    CreateModernButton ws, btnLeft, btnTop, btnWidth, btnHeight, _
        "5. Reduce Pass 3: SKILL.md", "modTrainingOrchestrator.RunReducePass3", RGB(128, 90, 213)
    btnTop = btnTop + btnHeight + 15
    
    CreateModernButton ws, btnLeft, btnTop, btnWidth, btnHeight, _
        "6. Save SKILL.md", "modTrainingOrchestrator.SaveSkillMd", RGB(128, 90, 213)
        
    ' --- GROUP 2: RUN REVIEW ---
    btnLeft = 340
    btnTop = 130
    
    Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, btnLeft, btnTop - 30, 300, 30)
    shp.Fill.Visible = msoFalse
    shp.Line.Visible = msoFalse
    shp.TextFrame2.TextRange.Text = "RUN REVIEW"
    shp.TextFrame2.TextRange.Font.Name = "Segoe UI"
    shp.TextFrame2.TextRange.Font.Size = 14
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(49, 130, 206) ' Blue
    
    CreateModernButton ws, btnLeft, btnTop, btnWidth, btnHeight, _
        "1. Select Persona for Review", "modDashboardUI.SetActivePersona", RGB(74, 85, 104)
    btnTop = btnTop + btnHeight + 15
    
    CreateModernButton ws, btnLeft, btnTop, btnWidth, btnHeight, _
        "2. Prepare Document for LLM", "modWordExport.ExportWordDocForLLM", RGB(49, 130, 206)
    btnTop = btnTop + btnHeight + 15
    
    CreateModernButton ws, btnLeft, btnTop, btnWidth, btnHeight, _
        "3. Apply LLM Edits to Word", "InputEditsIntoWord.ApplyWordSuggestionsFromJson", RGB(56, 161, 105) ' Green
    
    btnTop = btnTop + btnHeight + 40
    
    ' Button 4: Settings
    CreateModernButton ws, btnLeft, btnTop, 140, 35, _
        "Go to Config", "modDashboardUI.GoToConfigSheet", RGB(74, 85, 104) ' Gray
        
    ws.Range("A1").Select
End Sub

Public Sub SetActivePersona()
    Dim current As String
    Dim choice As String
    
    current = ConfigHelpers.GetConfigValue("ActivePersona", "")
    choice = InputBox("Enter the Persona Name:", "Set Active Persona", current)
    
    If Trim(choice) <> "" Then
        ConfigHelpers.SetConfigValue "ActivePersona", choice
        modPersonaRegistry.UpsertPersona choice
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
