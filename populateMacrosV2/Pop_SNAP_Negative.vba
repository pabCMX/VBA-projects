Attribute VB_Name = "Pop_SNAP_Negative"
' ============================================================================
' Pop_SNAP_Negative - SNAP Negative Schedule Population Module
' ============================================================================
' WHAT THIS MODULE DOES:
'   This module reads data from BIS delimited files and populates SNAP 
'   Negative (denial/termination/suspension) review schedule templates.
'   It's simpler than SNAP Positive since negative actions have fewer
'   data elements to populate.
'
' SNAP NEGATIVE ACTION TYPES:
'   - Denial (A): Application was rejected
'   - Termination (C): Benefits were ended
'   - Suspension (S): Benefits temporarily stopped
'
' BIS FILE STRUCTURE (Different from Positive):
'   The SNAP Negative BIS file has a simpler structure:
'   - Single worksheet with case-level data
'   - Review number in column A
'   - Action type in column C (A/C/S)
'   - Action date in column K (YYYYMMDD)
'   - Notice date in column S (YYYYMMDD)
'
' KEY FIELDS POPULATED:
'   - Date Assigned (today's date)
'   - Action Date (when the action took effect)
'   - Notice Date (when notice was sent, except for suspensions)
'   - Type of Negativity (1=Denial, 2=Termination, 3=Suspension)
'   - Narrative text box with action description
'
' DEPENDENCIES:
'   - Common_Utils: For helper functions
'   - Pop_Main: Sets up wb_bis and wb workbook references
'
' GLOBAL VARIABLES (from Pop_Main):
'   - wb_bis: BIS workbook
'   - wb: Output schedule workbook
'
' CHANGE LOG:
'   2026-01-03  Created from populate_snap_neg_delimited_mod.vba
'               Added Option Explicit, V2 comments, error handling
' ============================================================================

Option Explicit

' ============================================================================
' MAIN POPULATION FUNCTION
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: PopulateSNAPNegativeDelimited
' ----------------------------------------------------------------------------
' PURPOSE:
'   Populates a SNAP Negative schedule from BIS data.
'   This is simpler than SNAP Positive - fewer fields and no individual-level
'   data processing.
'
' PRE-CONDITIONS:
'   - wb_bis contains the opened BIS delimited file (single worksheet)
'   - wb contains the schedule workbook to populate
'
' POST-CONDITIONS:
'   - Schedule populated with action dates, type, and narrative
' ----------------------------------------------------------------------------
Public Sub PopulateSNAPNegativeDelimited()
    On Error GoTo ErrorHandler
    
    ' ========================================================================
    ' VARIABLE DECLARATIONS
    ' ========================================================================
    
    Dim wbSch As Workbook           ' Schedule workbook
    Dim wsBIS As Worksheet          ' BIS data worksheet
    Dim wsSchedule As Worksheet     ' Schedule worksheet
    Dim reviewNumber As String      ' Review number
    Dim caseRow As Range            ' Row in BIS for this review
    Dim maxrowBIS As Long           ' Last row in BIS sheet
    Dim actionType As String        ' A/C/S code
    Dim actionTypeName As String    ' "Denial"/"Termination"/"Suspension"
    Dim ws As Worksheet
    
    ' ========================================================================
    ' 1. SETUP - GET WORKBOOK REFERENCES
    ' ========================================================================
    
    Set wbSch = wb
    
    ' Find the review number from the schedule sheet name
    For Each ws In wbSch.Worksheets
        If Val(ws.Name) > 1000 Then
            reviewNumber = ws.Name
            Exit For
        End If
    Next ws
    
    If reviewNumber = "" Then
        MsgBox "Could not find schedule sheet (review number > 1000)", vbCritical
        Exit Sub
    End If
    
    ' Set worksheet references
    ' SNAP Negative BIS file has only one worksheet (index 1)
    Set wsBIS = wb_bis.Worksheets(1)
    Set wsSchedule = wbSch.Worksheets(reviewNumber)
    
    ' ========================================================================
    ' 2. POPULATE DATE ASSIGNED (TODAY'S DATE)
    ' ========================================================================
    
    ' Date Assigned is split across three cells: Month, Day, Year
    wsSchedule.Range("C16").Value = Application.WorksheetFunction.Text(Month(Date), "00")
    wsSchedule.Range("F16").Value = Application.WorksheetFunction.Text(Day(Date), "00")
    wsSchedule.Range("I16").Value = Year(Date)
    
    ' ========================================================================
    ' 3. FIND CASE IN BIS FILE
    ' ========================================================================
    
    maxrowBIS = wsBIS.Cells(wsBIS.Rows.Count, 1).End(xlUp).Row
    
    ' Review number is in column A for SNAP Negative
    With wsBIS.Range("A2:A" & maxrowBIS)
        Set caseRow = .Find(reviewNumber, LookIn:=xlValues)
    End With
    
    ' If case not found, exit (date assigned is already populated)
    If caseRow Is Nothing Then
        Exit Sub
    End If
    
    ' ========================================================================
    ' 4. POPULATE ACTION DATE
    ' ========================================================================
    
    ' Action date is in column K, format YYYYMMDD
    Dim actionDateStr As String
    actionDateStr = CStr(wsBIS.Range("K" & caseRow.Row).Value)
    
    If Len(actionDateStr) >= 8 Then
        wsSchedule.Range("S24").Value = Mid(actionDateStr, 5, 2)  ' Month
        wsSchedule.Range("V24").Value = Right(actionDateStr, 2)   ' Day
        wsSchedule.Range("Y24").Value = Left(actionDateStr, 4)    ' Year
    End If
    
    ' ========================================================================
    ' 5. POPULATE NOTICE DATE (Except for Suspensions)
    ' ========================================================================
    
    actionType = UCase(Trim(wsBIS.Range("C" & caseRow.Row).Value))
    
    ' Suspensions don't have a notice date (they happen immediately)
    If actionType <> "S" Then
        Dim noticeDateStr As String
        noticeDateStr = CStr(wsBIS.Range("S" & caseRow.Row).Value)
        
        If Len(noticeDateStr) >= 8 Then
            wsSchedule.Range("G24").Value = Mid(noticeDateStr, 5, 2)  ' Month
            wsSchedule.Range("J24").Value = Right(noticeDateStr, 2)   ' Day
            wsSchedule.Range("M24").Value = Left(noticeDateStr, 4)    ' Year
        End If
    End If
    
    ' ========================================================================
    ' 6. POPULATE TYPE OF NEGATIVITY
    ' ========================================================================
    
    Select Case actionType
        Case "A"
            wsSchedule.Range("AE24").Value = 1
            actionTypeName = "Denial"
        Case "C"
            wsSchedule.Range("AE24").Value = 2
            actionTypeName = "Termination"
        Case "S"
            wsSchedule.Range("AE24").Value = 3
            actionTypeName = "Suspension"
        Case Else
            actionTypeName = "Action"
    End Select
    
    ' ========================================================================
    ' 7. POPULATE NARRATIVE TEXT BOX
    ' ========================================================================
    
    ' Build the first sentence of the narrative
    Dim narrativeText As String
    narrativeText = "The action being reviewed is the SNAP " & actionTypeName & " of " & _
                    wsSchedule.Range("S24").Value & "/" & _
                    wsSchedule.Range("V24").Value & "/" & _
                    wsSchedule.Range("Y24").Value & "."
    
    ' Update the text box (Text Box 17 in the schedule)
    On Error Resume Next
    wsSchedule.Shapes("Text Box 17").TextFrame.Characters.Text = narrativeText
    On Error GoTo ErrorHandler
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in PopulateSNAPNegativeDelimited: " & Err.Description, vbCritical
    LogError "PopulateSNAPNegativeDelimited", Err.Number, Err.Description, "Review: " & reviewNumber
End Sub


