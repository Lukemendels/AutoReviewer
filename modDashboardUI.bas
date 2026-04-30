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
    
    ' Set column widths for alignment
    ws.Columns("A:A").ColumnWidth = 5
    ws.Columns("B:B").ColumnWidth = 60
    
    ' 4. Add Title
    Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, 30, 20, 400, 60)
    With shp
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        With .TextFrame2.TextRange
            .Text = "AutoReviewer Dashboard"
            .Font.Name = "Segoe UI"
            .Font.Size = 28
            .Font.Bold = msoTrue
            .Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        End With
    End With
    
    ' 5. Add Subtitle/Instructions
    Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, 30, 80, 500, 80)
    With shp
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        With .TextFrame2.TextRange
            .Text = "Welcome to the Paperclip API Workflow." & vbCrLf & _
                    "Extract style rules from your boss's edits, prepare new documents for review, and automatically apply the LLM's suggested track changes."
            .Font.Name = "Segoe UI"
            .Font.Size = 12
            .Font.Fill.ForeColor.RGB = RGB(160, 170, 180)
        End With
    End With
    
    ' Button Layout Params
    btnLeft = 30
    btnTop = 180
    btnWidth = 280
    btnHeight = 45
    
    ' Button 1: Extract Rules
    CreateModernButton ws, btnLeft, btnTop, btnWidth, btnHeight, _
        "1. Extract Boss's Rules", "modRuleExtractor.ExtractRulesAndPrompt", RGB(128, 90, 213) ' Purple
        
    btnTop = btnTop + btnHeight + 20
    
    ' Button 2: Prepare Document
    CreateModernButton ws, btnLeft, btnTop, btnWidth, btnHeight, _
        "2. Prepare Document for LLM", "modWordExport.ExportWordDocForLLM", RGB(49, 130, 206) ' Blue
        
    btnTop = btnTop + btnHeight + 20
    
    ' Button 3: Apply Edits
    CreateModernButton ws, btnLeft, btnTop, btnWidth, btnHeight, _
        "3. Apply LLM Edits to Word", "InputEditsIntoWord.ApplyWordSuggestionsFromJson", RGB(56, 161, 105) ' Green
        
    btnTop = btnTop + btnHeight + 40
    
    ' Button 4: Settings
    CreateModernButton ws, btnLeft, btnTop, 140, 35, _
        "Go to Config", "modDashboardUI.GoToConfigSheet", RGB(74, 85, 104) ' Gray
        
    ' Select A1
    ws.Range("A1").Select
    
    MsgBox "Dashboard created successfully!", vbInformation, "AutoReviewer"
End Sub

Private Sub CreateModernButton(ws As Worksheet, Left As Double, Top As Double, Width As Double, Height As Double, _
                               Text As String, MacroName As String, BgColor As Long)
    Dim shp As Shape
    
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, Left, Top, Width, Height)
    
    With shp
        ' Appearance
        .Fill.ForeColor.RGB = BgColor
        .Line.Visible = msoFalse
        .Adjustments.Item(1) = 0.15 ' border radius
        
        ' Shadow
        .Shadow.Type = msoShadow21
        .Shadow.ForeColor.RGB = RGB(0, 0, 0)
        .Shadow.Transparency = 0.6
        .Shadow.Size = 100
        .Shadow.Blur = 8
        .Shadow.OffsetX = 0
        .Shadow.OffsetY = 4
        
        ' Text properties
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
        
        ' Action mapping
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
