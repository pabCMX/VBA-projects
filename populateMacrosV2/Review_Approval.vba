Attribute VB_Name = "Review_Approval"
' ============================================================================
' Review_Approval - Supervisor and Clerical Approval Workflow
' ============================================================================
' WHAT THIS MODULE DOES:
'   This module handles the approval process when a QC review is completed.
'   It's used by examiners (not during population) to submit their finished
'   reviews for supervisor approval and then for final clerical processing.
'
' THE APPROVAL WORKFLOW:
'   1. Examiner completes the review schedule
'   2. Examiner clicks "Submit for Supervisor Approval" button
'      -> SupervisorApproval() runs, stamps username/date, saves to clerical folder
'   3. Supervisor reviews and approves
'      -> May click "Supervisor Approval" to add their stamp
'   4. Clerical staff processes the final review
'      -> ClericalApproval() runs, stamps username/date, saves to QCMIS folder
'
' WHERE FILES GO:
'   The approval functions save copies of the workbook to specific network
'   folders based on:
'   - Program type (TANF, SNAP, MA, GA)
'   - Status (Clean, Error, Drop)
'   - Region/Examiner assignment
'
' KEY FUNCTIONS:
'   - SupervisorApproval() : Main supervisor approval (stamps and routes)
'   - ClericalApproval()   : Final clerical processing
'   - SuprWorkbook()       : Supervisor stamp on workbook sheet only
'   - tracking_supr()      : Simple tracking timestamp
'
' IMPROVEMENTS FROM ORIGINAL:
'   - Uses Common_Utils for network drive detection (was duplicated)
'   - Uses Common_Utils for status folder determination (was duplicated)
'   - Uses Config_Settings for paths and constants
'   - Proper error handling throughout
'   - Option Explicit to catch typos
'
' IMPORTANT CELL REFERENCES:
'   Each program has different cells for approval stamps:
'   - TANF/GA: AL5 (supervisor), AL6 (clerical)
'   - SNAP+:   AH2 (supervisor), AH13 (clerical)
'   - SNAP-:   AC17 (supervisor), AK15 (clerical)
'   - MA+:     AL5 (supervisor), AL6 (clerical)
'   - MA-:     AB2 (supervisor), AB3 (clerical)
'
' CHANGE LOG:
'   2026-01-02  Refactored from Module1 - extracted shared code to Common_Utils,
'               added Option Explicit, improved error handling, V2 comments
' ============================================================================

Option Explicit


' ============================================================================
' SUPERVISOR APPROVAL
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: SupervisorApproval
' ----------------------------------------------------------------------------
' PURPOSE:
'   Called when the supervisor approves a completed review. This function:
'   1. Records the supervisor's username and date on the schedule
'   2. Records the same info on the program workbook sheet
'   3. Determines the review status (Clean/Error/Drop)
'   4. Saves the workbook to the appropriate clerical folder
'
' ENTRY POINT:
'   This is attached to the "Supervisor Approval" button on the schedule.
'
' PRE-CONDITIONS:
'   - Network drive must be mapped
'   - Review must be complete (all required fields filled)
'
' POST-CONDITIONS:
'   - Approval timestamp added to schedule
'   - Workbook saved to clerical folder
'   - For SNAP+: May show timeliness memo reminders
' ----------------------------------------------------------------------------
Public Sub SupervisorApproval()
    On Error GoTo ErrorHandler
    
    ' ========================================================================
    ' 1. VARIABLE DECLARATIONS
    ' ========================================================================
    Dim thisWB As Workbook
    Dim thisWS As Worksheet
    Dim dqcPath As String
    Dim clericalPath As String
    Dim programFolder As String
    Dim statusFolder As String
    Dim examNum As Long
    Dim programType As ProgramType
    Dim reviewStatus As StatusFolder
    Dim savePath As String
    
    Set thisWB = ActiveWorkbook
    Set thisWS = ActiveSheet
    
    ' ========================================================================
    ' 2. VALIDATE NETWORK PATH
    ' ========================================================================
    ' Get the network drive path using our centralized function
    ' This replaces 20+ lines of duplicated code from the original
    
    dqcPath = GetDQCDriveLetter()
    If dqcPath = "" Then
        MsgBox "Network Drive to Examiner Files is NOT correct" & vbCrLf & _
               "Contact Valerie or Nicole", vbCritical, "Network Error"
        Exit Sub
    End If
    
    ' Build paths to clerical folders
    ' Different regions may use different folders - adjust as needed
    clericalPath = dqcPath & "HBG Clerical Schedules\"
    
    If Not PathExists(clericalPath) Then
        MsgBox "Path to Clerical Folder: " & clericalPath & " does NOT exist!!" & vbCrLf & _
               "Contact Valerie or Nicole", vbCritical, "Path Error"
        Exit Sub
    End If
    
    ' ========================================================================
    ' 3. DETERMINE PROGRAM TYPE
    ' ========================================================================
    ' Use the sheet name to determine which program this review is for
    
    programType = GetProgramFromSheetName(thisWS.Name)
    
    ' ========================================================================
    ' 4. PROGRAM-SPECIFIC PROCESSING
    ' ========================================================================
    ' Each program has different:
    '   - Cells for approval stamps
    '   - Workbook sheets to update
    '   - Logic for determining Clean/Error/Drop status
    
    Select Case programType
    
        ' ----------------------------------------------------------------
        ' TANF REVIEWS
        ' ----------------------------------------------------------------
        Case PROG_TANF
            ' Add supervisor timestamp to schedule
            thisWS.Range("AL5").Value = GetUserTimestamp()
            
            ' Add timestamp to TANF Workbook sheet
            On Error Resume Next
            thisWB.Worksheets("TANF Workbook").Range("G41").Value = Environ("USERNAME")
            thisWB.Worksheets("TANF Workbook").Range("G44").Value = Date
            On Error GoTo ErrorHandler
            
            ' Determine status using centralized function
            reviewStatus = DetermineStatusFolder(PROG_TANF, thisWS)
            statusFolder = StatusFolderToString(reviewStatus)
            
            ' Save to clerical folder
            thisWB.Save
            programFolder = "TANF"
            savePath = clericalPath & programFolder & "\" & statusFolder & "\" & thisWB.Name
            thisWB.SaveAs Filename:=savePath
            
        ' ----------------------------------------------------------------
        ' GA REVIEWS
        ' ----------------------------------------------------------------
        Case PROG_GA
            ' Add supervisor timestamp
            thisWS.Range("AL5").Value = GetUserTimestamp()
            
            ' Add timestamp to GA Workbook sheet
            On Error Resume Next
            thisWB.Worksheets("GA Workbook").Range("G41").Value = Environ("USERNAME")
            thisWB.Worksheets("GA Workbook").Range("G44").Value = Date
            On Error GoTo ErrorHandler
            
            ' Determine status
            reviewStatus = DetermineStatusFolder(PROG_GA, thisWS)
            statusFolder = StatusFolderToString(reviewStatus)
            
            ' Save - GA may go to different folder
            thisWB.Save
            programFolder = "GA"
            savePath = clericalPath & programFolder & "\" & statusFolder & "\" & thisWB.Name
            thisWB.SaveAs Filename:=savePath
            
        ' ----------------------------------------------------------------
        ' MA POSITIVE REVIEWS
        ' ----------------------------------------------------------------
        Case PROG_MA_POS
            ' Add supervisor timestamp
            thisWS.Range("AL5").Value = GetUserTimestamp()
            
            ' Add timestamp to MA Workbook sheet
            On Error Resume Next
            thisWB.Worksheets("MA Workbook").Range("F43").Value = Environ("USERNAME")
            thisWB.Worksheets("MA Workbook").Range("F45").Value = Date
            On Error GoTo ErrorHandler
            
            ' Determine status
            reviewStatus = DetermineStatusFolder(PROG_MA_POS, thisWS)
            statusFolder = StatusFolderToString(reviewStatus)
            
            ' Save
            thisWB.Save
            programFolder = "MA Positive"
            savePath = clericalPath & programFolder & "\" & statusFolder & "\" & thisWB.Name
            thisWB.SaveAs Filename:=savePath
            
        ' ----------------------------------------------------------------
        ' SNAP POSITIVE REVIEWS
        ' ----------------------------------------------------------------
        Case PROG_SNAP_POS
            ' Special validation for SNAP+: Check expedited block
            If thisWS.Range("B157").Value = "" Then
                thisWS.Range("B157").Select
                MsgBox "The SNAP Expedited block (row 3, 1st block) in Section 7 must be filled in." & vbCrLf & _
                       "Please fill in the block and click the Supervisor Approval button again.", _
                       vbExclamation, "Required Field Missing"
                Exit Sub
            End If
            
            ' Add supervisor timestamp
            thisWS.Range("AH2").Value = GetUserTimestamp()
            
            ' Add timestamp to FS Workbook sheet
            On Error Resume Next
            thisWB.Worksheets("FS Workbook").Range("G42").Value = Environ("USERNAME")
            thisWB.Worksheets("FS Workbook").Range("G44").Value = Date
            On Error GoTo ErrorHandler
            
            ' Determine status
            reviewStatus = DetermineStatusFolder(PROG_SNAP_POS, thisWS)
            statusFolder = StatusFolderToString(reviewStatus)
            
            ' Save
            thisWB.Save
            programFolder = "SNAP Positive"
            savePath = clericalPath & programFolder & "\" & statusFolder & "\" & thisWB.Name
            thisWB.SaveAs Filename:=savePath
            
            ' Timeliness memo reminders for specific finding codes
            Call CheckTimelinessReminders(thisWS)
            
        ' ----------------------------------------------------------------
        ' SNAP NEGATIVE REVIEWS
        ' ----------------------------------------------------------------
        Case PROG_SNAP_NEG
            ' Add supervisor timestamp and date components
            thisWS.Range("AC17").Value = GetUserTimestamp()
            
            ' Parse date into separate cells (legacy format requirement)
            Dim dateParts() As String
            dateParts = Split(CStr(Date), "/")
            thisWS.Range("AA16").Value = dateParts(LBound(dateParts))       ' Month
            thisWS.Range("AD16").Value = dateParts(LBound(dateParts) + 1)   ' Day
            thisWS.Range("AG16").Value = dateParts(UBound(dateParts))       ' Year
            
            ' Determine status
            reviewStatus = DetermineStatusFolder(PROG_SNAP_NEG, thisWS)
            statusFolder = StatusFolderToString(reviewStatus)
            
            ' Save
            thisWB.Save
            programFolder = "SNAP Negative"
            savePath = clericalPath & programFolder & "\" & statusFolder & "\" & thisWB.Name
            thisWB.SaveAs Filename:=savePath
            
        ' ----------------------------------------------------------------
        ' MA NEGATIVE REVIEWS
        ' ----------------------------------------------------------------
        Case PROG_MA_NEG, PROG_MA_PE
            ' Add supervisor timestamp
            thisWS.Range("AB2").Value = GetUserTimestamp()
            
            ' Determine status
            reviewStatus = DetermineStatusFolder(PROG_MA_NEG, thisWS)
            statusFolder = StatusFolderToString(reviewStatus)
            
            ' Save
            thisWB.Save
            programFolder = "MA Negative"
            savePath = clericalPath & programFolder & "\" & statusFolder & "\" & thisWB.Name
            thisWB.SaveAs Filename:=savePath
            
        Case Else
            MsgBox "Unknown program type for sheet: " & thisWS.Name, _
                   vbExclamation, "Unknown Program"
            Exit Sub
    End Select
    
    ' Return cursor to cell A1
    thisWS.Range("A1").Select
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Supervisor Approval Error"
    LogError "SupervisorApproval", Err.Number, Err.Description, "Sheet: " & thisWS.Name
End Sub


' ============================================================================
' CLERICAL APPROVAL
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: ClericalApproval
' ----------------------------------------------------------------------------
' PURPOSE:
'   Called when clerical staff processes a completed review. This function:
'   1. Records the clerical staff's username and date on the schedule
'   2. Saves the workbook to the QCMIS folder for final archiving
'
' ENTRY POINT:
'   Attached to the "Clerical Approval" button on the schedule.
'
' POST-CONDITIONS:
'   - Clerical timestamp added to schedule
'   - Workbook saved to QCMIS folder
' ----------------------------------------------------------------------------
Public Sub ClericalApproval()
    On Error GoTo ErrorHandler
    
    Dim thisWB As Workbook
    Dim thisWS As Worksheet
    Dim dqcPath As String
    Dim qcmisPath As String
    Dim programFolder As String
    Dim statusFolder As String
    Dim programType As ProgramType
    Dim reviewStatus As StatusFolder
    Dim savePath As String
    Dim dateParts() As String
    
    Set thisWB = ActiveWorkbook
    Set thisWS = ActiveSheet
    
    ' ========================================================================
    ' VALIDATE NETWORK PATH
    ' ========================================================================
    dqcPath = GetDQCDriveLetter()
    If dqcPath = "" Then
        MsgBox "Network Drive to Examiner Files is NOT correct" & vbCrLf & _
               "Contact Valerie or Nicole", vbCritical, "Network Error"
        Exit Sub
    End If
    
    qcmisPath = dqcPath & "QCMIS Schedules\"
    
    If Not PathExists(qcmisPath) Then
        MsgBox "Path to QCMIS Folder: " & qcmisPath & " does NOT exist!!" & vbCrLf & _
               "Contact Valerie or Nicole", vbCritical, "Path Error"
        Exit Sub
    End If
    
    ' ========================================================================
    ' DETERMINE PROGRAM TYPE AND ADD TIMESTAMPS
    ' ========================================================================
    programType = GetProgramFromSheetName(thisWS.Name)
    
    Select Case programType
    
        Case PROG_TANF, PROG_GA
            thisWS.Range("AL6").Value = GetUserTimestamp()
            
            reviewStatus = DetermineStatusFolder(programType, thisWS)
            statusFolder = StatusFolderToString(reviewStatus)
            
            thisWB.Save
            
            ' Determine program folder based on review number
            If Left(thisWS.Range("A10").Value, 1) = "1" Then
                programFolder = "TANF"
            Else
                programFolder = "GA"
            End If
            
            savePath = qcmisPath & programFolder & "\" & statusFolder & "\" & thisWB.Name
            thisWB.SaveAs Filename:=savePath
            
        Case PROG_MA_POS
            thisWS.Range("AL6").Value = GetUserTimestamp()
            
            reviewStatus = DetermineStatusFolder(PROG_MA_POS, thisWS)
            statusFolder = StatusFolderToString(reviewStatus)
            
            thisWB.Save
            programFolder = "MA Positive"
            savePath = qcmisPath & programFolder & "\" & statusFolder & "\" & thisWB.Name
            thisWB.SaveAs Filename:=savePath
            
        Case PROG_SNAP_POS
            thisWS.Range("AH13").Value = GetUserTimestamp()
            
            reviewStatus = DetermineStatusFolder(PROG_SNAP_POS, thisWS)
            statusFolder = StatusFolderToString(reviewStatus)
            
            thisWB.Save
            programFolder = "SNAP Positive"
            savePath = qcmisPath & programFolder & "\" & statusFolder & "\" & thisWB.Name
            thisWB.SaveAs Filename:=savePath
            
        Case PROG_SNAP_NEG
            thisWS.Range("AK15").Value = Environ("USERNAME")
            
            ' Parse date components
            dateParts = Split(CStr(Date), "/")
            thisWS.Range("AA16").Value = dateParts(LBound(dateParts))
            thisWS.Range("AD16").Value = dateParts(LBound(dateParts) + 1)
            thisWS.Range("AG16").Value = dateParts(UBound(dateParts))
            
            reviewStatus = DetermineStatusFolder(PROG_SNAP_NEG, thisWS)
            statusFolder = StatusFolderToString(reviewStatus)
            
            thisWB.Save
            programFolder = "SNAP Negative"
            savePath = qcmisPath & programFolder & "\" & statusFolder & "\" & thisWB.Name
            thisWB.SaveAs Filename:=savePath
            
        Case PROG_MA_NEG, PROG_MA_PE
            thisWS.Range("AB3").Value = GetUserTimestamp()
            
            ' MA Negative uses different status determination
            Dim maNegStatus As String
            If thisWS.Range("M56").Value <> 0 Then
                statusFolder = "Drop"
            ElseIf UCase(thisWS.Range("AC1").Value) = "ERROR" Then
                statusFolder = "Error"
            Else
                statusFolder = "Clean"
            End If
            
            thisWB.Save
            programFolder = "MA Negative"
            savePath = qcmisPath & programFolder & "\" & statusFolder & "\" & thisWB.Name
            thisWB.SaveAs Filename:=savePath
            
        Case Else
            MsgBox "Unknown program type for sheet: " & thisWS.Name, _
                   vbExclamation, "Unknown Program"
            Exit Sub
    End Select
    
    thisWS.Range("A1").Select
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Clerical Approval Error"
    LogError "ClericalApproval", Err.Number, Err.Description, "Sheet: " & thisWS.Name
End Sub


' ============================================================================
' HELPER FUNCTIONS
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: tracking_supr
' ----------------------------------------------------------------------------
' PURPOSE:
'   Simple function to add a supervisor tracking timestamp to cell M13.
'   Used for internal tracking purposes.
' ----------------------------------------------------------------------------
Public Sub tracking_supr()
    On Error Resume Next
    Range("M13").Value = GetUserTimestamp()
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' Sub: SuprWorkbook
' ----------------------------------------------------------------------------
' PURPOSE:
'   Adds supervisor stamp to the program workbook sheet (not the schedule).
'   This is separate from the main approval so supervisors can stamp the
'   workbook without triggering the full approval workflow.
' ----------------------------------------------------------------------------
Public Sub SuprWorkbook()
    On Error Resume Next
    
    Select Case ActiveSheet.Name
        Case "TANF Workbook", "GA Workbook"
            Range("G41").Value = Environ("USERNAME")
            Range("G44").Value = Date
        Case "FS Workbook"
            Range("G42").Value = Environ("USERNAME")
            Range("G44").Value = Date
    End Select
    
    Range("A1").Select
    
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' Sub: CheckTimelinessReminders
' ----------------------------------------------------------------------------
' PURPOSE:
'   For SNAP Positive reviews, checks if the finding codes indicate a need
'   for a timeliness memo and displays a reminder message.
'
' PARAMETERS:
'   ws - The worksheet to check
'
' FINDING CODES THAT TRIGGER REMINDERS:
'   - C149 = 2 OR K149 = 11, 12, or 13: Findings Memo
'   - K149 = 24, 25, 26, or 27: Info Memo
' ----------------------------------------------------------------------------
Private Sub CheckTimelinessReminders(ByRef ws As Worksheet)
    On Error Resume Next
    
    ' Check for Timeliness Findings Memo
    If ws.Range("C149").Value = 2 Or _
       ws.Range("K149").Value = 11 Or _
       ws.Range("K149").Value = 12 Or _
       ws.Range("K149").Value = 13 Then
        MsgBox "Check for SNAP Timeliness Findings Memo", vbCritical, "Timeliness"
    End If
    
    ' Check for Timeliness Info Memo
    If ws.Range("K149").Value = 24 Or _
       ws.Range("K149").Value = 25 Or _
       ws.Range("K149").Value = 26 Or _
       ws.Range("K149").Value = 27 Then
        MsgBox "Check for SNAP Timeliness Info Memo", vbCritical, "Timeliness"
    End If
    
    On Error GoTo 0
End Sub


