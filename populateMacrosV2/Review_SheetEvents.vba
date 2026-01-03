Attribute VB_Name = "Review_SheetEvents"
' ============================================================================
' Review_SheetEvents - Worksheet_Change Event Handlers for Examiner Schedules
' ============================================================================
' WHAT THIS MODULE DOES:
'   This module consolidates all the Worksheet_Change event handlers that were
'   originally scattered across individual sheet modules (Sheet4, Sheet8,
'   Sheet12, Sheet14, Sheet18, Sheet20, Sheet22, Sheet25, Sheet27).
'
'   These event handlers provide real-time validation and auto-fill features
'   as examiners enter data into the schedule.
'
' IMPORTANT: HOW SHEET EVENTS WORK IN VBA
'   Worksheet_Change events must be placed in the SHEET'S code module
'   (not a standard module). During the repopulate process, this consolidated
'   logic should be copied to the appropriate sheet code modules.
'
' KEY PATTERNS CONSOLIDATED:
'
'   1. VEHICLE BOX AUTO-FILL (Sheet8, Sheet14)
'      When vehicle value = 1, fill adjacent cell with "-"
'
'   2. COMPUTATION CELL CLEARING (Sheet4, Sheet18, Sheet22, Sheet27)
'      When source cell is cleared, clear dependent calculated cell
'
'   3. DROP CODE SYNCING (Sheet12)
'      Sync drop code between schedule and workbook sheets
'
'   4. SUA USAGE/PRORATION LOGIC (Sheet14)
'      When SUA usage = 1, set proration to "-"
'
'   5. ALLOTMENT ADJUSTMENT (Sheet14)
'      Auto-fill based on adjustment code selection
'
'   6. DYNAMIC ELEMENT CODE VALIDATION (Sheet12, Sheet14, Sheet20, Sheet25)
'      When an element code is entered:
'      - Look up valid nature codes from reference range
'      - Create data validation dropdown
'      - Add comment with code descriptions
'
'   7. NARRATIVE TEMPLATE AUTO-FILL (Sheet25)
'      When dropdown selection is made, fill text box with template
'
' HOW TO USE THIS MODULE:
'   The event handlers in this module are TEMPLATES. During population:
'   1. Copy the appropriate Worksheet_Change handler to the sheet's code module
'   2. Adjust cell references for the specific sheet/program
'
'   For examiners, these events fire automatically when they edit cells.
'
' CHANGE LOG:
'   2026-01-03  Consolidated from 9 sheet modules
'               Added V2 comments, modular structure
' ============================================================================

Option Explicit

' ============================================================================
' PATTERN 1: VEHICLE BOX AUTO-FILL
' ============================================================================
' PURPOSE:
'   When the vehicle checkbox/value is set to 1, automatically fill the
'   adjacent explanation cell with "-" (indicating N/A).
'
' USED IN: SNAP Positive (Sheet8, Sheet14)
' CELLS: P62 -> V62 (typical)
' ============================================================================
Public Sub HandleVehicleBox(ByVal Target As Range, ByVal ws As Worksheet)
    On Error Resume Next
    
    ' Check if target is the vehicle value cell (row 62, column P)
    If Target.Row = 62 And Target.Column = 16 Then
        If Target.Value = 1 Then
            ws.Range("V62").Value = "-"
        Else
            ws.Range("V62").Value = ""
        End If
    End If
    
    On Error GoTo 0
End Sub


' ============================================================================
' PATTERN 2: COMPUTATION CELL CLEARING
' ============================================================================
' PURPOSE:
'   When a source cell is cleared (or set to specific values), clear the
'   dependent computed cell. This prevents stale calculations.
'
' USED IN: Computation sheets (Sheet4, Sheet18, Sheet22, Sheet27)
' ============================================================================
Public Sub HandleComputationClear(ByVal Target As Range, ByVal ws As Worksheet)
    On Error Resume Next
    
    ' Example: When column D is cleared, clear column E for same row
    ' Actual mappings vary by sheet - this is a template
    
    ' Pattern for TANF/SNAP computation sheets:
    ' If source value cell is cleared, clear the result cell
    If Target.Column = 4 Then  ' Column D
        If IsEmpty(Target.Value) Or Target.Value = "" Then
            ws.Cells(Target.Row, 5).Value = ""  ' Column E
        End If
    End If
    
    On Error GoTo 0
End Sub


' ============================================================================
' PATTERN 3: DROP CODE SYNCING
' ============================================================================
' PURPOSE:
'   Keeps the drop code synchronized between the main schedule and the
'   workbook sheet. When disposition changes on one, update the other.
'
' USED IN: TANF (Sheet12)
' LOGIC:
'   - If disposition (AI10) is not "1" (not clean), copy to workbook drop cell
'   - If disposition is "1" (clean), clear workbook drop cell
' ============================================================================
Public Sub HandleDropCodeSync(ByVal Target As Range, ByVal wsSchedule As Worksheet, _
                               ByVal wsWorkbook As Worksheet)
    On Error Resume Next
    
    ' Check if target is the disposition cell (AI10)
    If Target.Address = "$AI$10" Then
        If wsSchedule.Range("AI10").Value <> "1" Then
            ' Copy drop code to workbook
            wsWorkbook.Range("AE45").Value = wsSchedule.Range("AI10").Value
        Else
            ' Clean case - clear workbook drop
            wsWorkbook.Range("AE45").Value = ""
        End If
    End If
    
    On Error GoTo 0
End Sub


' ============================================================================
' PATTERN 4: SUA USAGE / PRORATION LOGIC
' ============================================================================
' PURPOSE:
'   Enforces business rules for Standard Utility Allowance:
'   - If SUA Usage (W82) = 1 and disposition = 1, then Proration = "-"
'   - If SUA Usage > 1, clear Proration
'
' USED IN: SNAP Positive (Sheet14)
' ============================================================================
Public Sub HandleSUALogic(ByVal Target As Range, ByVal ws As Worksheet)
    On Error Resume Next
    
    ' Check if disposition cell changed
    If Target.Row = 22 And Target.Column = 3 Then
        If Target.Value = 1 And ws.Range("W82").Value = "1" Then
            ws.Range("AA82").Value = "-"
        End If
    End If
    
    ' Check if SUA Usage cell changed
    If Target.Row = 82 And Target.Column = 23 Then
        If Target.Value = 1 And ws.Range("C22").Value = "1" Then
            ws.Range("AA82").Value = "-"
        ElseIf Val(ws.Range("W82").Value) > 1 Then
            ws.Range("AA82").Value = ""
        End If
    End If
    
    On Error GoTo 0
End Sub


' ============================================================================
' PATTERN 5: ALLOTMENT ADJUSTMENT
' ============================================================================
' PURPOSE:
'   Auto-fills the allotment adjustment explanation based on code selected:
'   - Code 1: Fill with "-" (no adjustment needed)
'   - Code 2 or 3: Fill with " " (adjustment explanation needed)
'
' USED IN: SNAP Positive (Sheet14)
' ============================================================================
Public Sub HandleAllotmentAdjustment(ByVal Target As Range, ByVal ws As Worksheet)
    On Error Resume Next
    
    ' Check if allotment adjustment code cell changed (AB50)
    If Target.Row = 50 And Target.Column = 28 Then
        Select Case Val(ws.Range("AB50").Value)
            Case 1
                ws.Range("AI50").Value = "-"
            Case 2, 3
                ws.Range("AI50").Value = " "
        End Select
    End If
    
    On Error GoTo 0
End Sub


' ============================================================================
' PATTERN 6: DYNAMIC ELEMENT CODE VALIDATION
' ============================================================================
' PURPOSE:
'   When an element code is entered, this pattern:
'   1. Looks up the valid nature codes for that element
'   2. Creates a data validation dropdown for the nature code cell
'   3. Adds a cell comment with code descriptions
'
'   This is the most complex pattern and appears in 4+ sheet modules.
'
' HOW IT WORKS:
'   - Element codes are stored in a reference column (e.g., BT)
'   - Nature codes are in an adjacent column (e.g., BU)
'   - Descriptions are in another column (e.g., BV)
'   - When user enters an element code, we find matching rows and create
'     validation for the nature code cell
'
' PARAMETERS:
'   Target - The cell that was changed
'   ws - The worksheet
'   elementCodeCol - Column containing element codes (e.g., "BT")
'   natureCodeCol - Column containing nature codes (e.g., "BU")
'   descriptionCol - Column containing descriptions (e.g., "BV")
'   targetCells - Array of row/column pairs that trigger this logic
'   natureColOffset - Offset from element cell to nature cell
' ============================================================================
Public Sub HandleElementCodeValidation(ByVal Target As Range, _
                                        ByVal ws As Worksheet, _
                                        ByVal elementCodeCol As String, _
                                        ByVal natureCodeCol As String, _
                                        ByVal descriptionCol As String, _
                                        ByVal natureColNumber As Long)
    On Error Resume Next
    
    ' Skip if target is empty
    If IsEmpty(Target.Value) Then Exit Sub
    
    Dim startRow As Long, endRow As Long, lastRow As Long
    Dim i As Long, j As Long
    Dim tempStr As String
    
    Application.EnableEvents = False
    
    ' Find the range of rows matching the entered element code
    startRow = 0
    endRow = 0
    lastRow = ws.Range(elementCodeCol & ws.Rows.Count).End(xlUp).Row
    
    For i = 2 To lastRow
        If Target.Value = ws.Range(elementCodeCol & i).Value And startRow = 0 Then
            startRow = i
        ElseIf Target.Value <> ws.Range(elementCodeCol & i).Value And startRow <> 0 Then
            endRow = i - 1
            Exit For
        End If
    Next i
    
    If startRow = 0 Then
        Application.EnableEvents = True
        Exit Sub
    End If
    
    If endRow = 0 Then endRow = lastRow
    
    ' Build comment string from descriptions
    tempStr = ""
    For j = startRow To endRow
        tempStr = tempStr & ws.Range(descriptionCol & j).Value & vbCrLf
    Next j
    
    ' Add comment to the nature code cell
    On Error Resume Next
    ws.Cells(Target.Row, natureColNumber).Comment.Delete
    ws.Cells(Target.Row, natureColNumber).AddComment
    ws.Cells(Target.Row, natureColNumber).Comment.Text tempStr
    On Error GoTo 0
    
    ' Create data validation for nature code cell
    tempStr = "=$" & natureCodeCol & "$" & startRow & ":$" & natureCodeCol & "$" & endRow
    
    With ws.Cells(Target.Row, natureColNumber).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:=tempStr
        .IgnoreBlank = True
        .InCellDropdown = False
        .ErrorTitle = "Nature"
        .ErrorMessage = "Please enter a valid Nature code. Refer to comment for descriptions."
        .ShowInput = True
        .ShowError = True
    End With
    
    Application.EnableEvents = True
    On Error GoTo 0
End Sub


' ============================================================================
' PATTERN 7: NARRATIVE TEMPLATE AUTO-FILL
' ============================================================================
' PURPOSE:
'   When a dropdown selection is made, fill the associated text box with
'   a narrative template that the examiner can then customize.
'
' USED IN: SNAP Negative (Sheet25)
' ============================================================================
Public Sub HandleNarrativeTemplate(ByVal Target As Range, ByVal ws As Worksheet)
    On Error Resume Next
    
    ' Example: When dropdown in specific cell changes, update text box
    ' Actual implementation depends on the specific narrative mappings
    
    ' Template pattern - customize for specific dropdowns
    ' If Target.Address = "$A$10" Then
    '     Select Case Target.Value
    '         Case "Option 1"
    '             ws.Shapes("Text Box 1").TextFrame.Characters.Text = "Template 1 text..."
    '         Case "Option 2"
    '             ws.Shapes("Text Box 1").TextFrame.Characters.Text = "Template 2 text..."
    '     End Select
    ' End If
    
    On Error GoTo 0
End Sub


' ============================================================================
' CONSOLIDATED WORKSHEET_CHANGE TEMPLATE FOR SNAP POSITIVE
' ============================================================================
' PURPOSE:
'   This is a TEMPLATE for the Worksheet_Change event handler that should
'   be placed in the SNAP Positive schedule sheet's code module.
'
' TO USE:
'   1. Copy this sub to the sheet's code module
'   2. Rename to "Worksheet_Change"
'   3. Adjust cell references if needed
' ============================================================================
Public Sub Template_SNAPPositive_Change(ByVal Target As Range)
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = Target.Worksheet
    
    Application.EnableEvents = True
    
    ' Vehicle Box auto-fill
    Call HandleVehicleBox(Target, ws)
    
    ' Allotment Adjustment
    Call HandleAllotmentAdjustment(Target, ws)
    
    ' SUA Logic
    Call HandleSUALogic(Target, ws)
    
    ' Element Code Validation (for error element cells)
    ' Rows 29, 31, 33, 35, 37, 39, 41, 43 - Column B
    If Target.Column = 2 Then
        Select Case Target.Row
            Case 29, 31, 33, 35, 37, 39, 41, 43
                Call HandleElementCodeValidation(Target, ws, "BT", "BU", "BV", 7)
        End Select
    End If
    
    Application.EnableEvents = True
    On Error GoTo 0
End Sub


' ============================================================================
' CONSOLIDATED WORKSHEET_CHANGE TEMPLATE FOR TANF
' ============================================================================
Public Sub Template_TANF_Change(ByVal Target As Range)
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim wsWorkbook As Worksheet
    
    Set ws = Target.Worksheet
    On Error Resume Next
    Set wsWorkbook = ws.Parent.Worksheets("TANF Workbook")
    On Error GoTo 0
    
    Application.EnableEvents = True
    
    ' Drop Code Sync
    If Not wsWorkbook Is Nothing Then
        Call HandleDropCodeSync(Target, ws, wsWorkbook)
    End If
    
    ' Element Code Validation (rows 61, 63, 65, 67 - Column J)
    If Target.Column = 10 Then
        Select Case Target.Row
            Case 61, 63, 65, 67
                Call HandleElementCodeValidation(Target, ws, "BP", "BQ", "BR", 15)
        End Select
    End If
    
    Application.EnableEvents = True
    On Error GoTo 0
End Sub


