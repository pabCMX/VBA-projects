Attribute VB_Name = "Review_Drop"
' ============================================================================
' Review_Drop - Schedule Clearing for Dropped Reviews
' ============================================================================
' WHAT THIS MODULE DOES:
'   When a review is "dropped" (not completed for various reasons like
'   the case being closed, unable to locate client, etc.), the schedule
'   needs to be cleared of any partial data. This module provides functions
'   to clear the appropriate sections based on program type.
'
' WHEN A REVIEW IS DROPPED:
'   - The examiner couldn't complete the review for a valid reason
'   - Disposition code is set to a value > 1
'   - Most data fields should be cleared to avoid confusion
'   - The review is filed in the "Drop" folder instead of Clean/Error
'
' COMMON DROP REASONS:
'   - Case closed before review could be completed
'   - Client could not be located
'   - Client refused to cooperate
'   - Duplicate case
'   - Out of scope for review period
'
' FUNCTIONS:
'   - ClearScheduleForDrop() : Main entry point, routes to program-specific
'   - ClearTANFSchedule()    : Clears TANF-specific fields
'   - ClearSNAPPosSchedule() : Clears SNAP Positive fields
'   - ClearMAPosSchedule()   : Clears MA Positive fields
'
' WHAT GETS CLEARED:
'   - Person-level data (names, SSNs, demographics)
'   - Income and resource amounts
'   - Error finding details
'   - Computation results
'   - But NOT: Review number, case number, disposition code, drop reason
'
' CHANGE LOG:
'   2026-01-02  Refactored from Drop.vba - added Option Explicit,
'               centralized program detection, V2 commenting style
' ============================================================================

Option Explicit


' ============================================================================
' MAIN ENTRY POINT
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: ClearScheduleForDrop
' ----------------------------------------------------------------------------
' PURPOSE:
'   Main entry point for clearing a schedule when a review is dropped.
'   Determines the program type and calls the appropriate clearing function.
'
' ENTRY POINT:
'   Attached to "Clear for Drop" button on schedule (if present).
' ----------------------------------------------------------------------------
Public Sub ClearScheduleForDrop()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim programType As ProgramType
    Dim confirmResult As VbMsgBoxResult
    
    Set ws = ActiveSheet
    
    ' Confirm with user - this action clears data
    confirmResult = MsgBox("This will clear all data from the schedule for a dropped review." & vbCrLf & vbCrLf & _
                          "Are you sure you want to continue?", _
                          vbQuestion + vbYesNo, "Confirm Clear")
    
    If confirmResult <> vbYes Then
        MsgBox "Clear cancelled.", vbInformation
        Exit Sub
    End If
    
    ' Determine program type
    programType = GetProgramFromSheetName(ws.Name)
    
    ' Route to appropriate clearing function
    Select Case programType
        Case PROG_TANF, PROG_GA
            Call ClearTANFSchedule(ws)
        Case PROG_SNAP_POS
            Call ClearSNAPPosSchedule(ws)
        Case PROG_MA_POS
            Call ClearMAPosSchedule(ws)
        Case Else
            MsgBox "Clear function not available for this program type.", _
                   vbExclamation, "Not Supported"
            Exit Sub
    End Select
    
    MsgBox "Schedule cleared for dropped review.", vbInformation, "Clear Complete"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error clearing schedule: " & Err.Description, vbCritical, "Error"
    LogError "ClearScheduleForDrop", Err.Number, Err.Description, ""
End Sub


' ============================================================================
' PROGRAM-SPECIFIC CLEARING FUNCTIONS
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: ClearTANFSchedule
' ----------------------------------------------------------------------------
' PURPOSE:
'   Clears data from a TANF schedule for a dropped review.
'   Also works for GA schedules since they have the same layout.
'
' PARAMETERS:
'   ws - The worksheet to clear
'
' RANGES CLEARED:
'   - Person information section
'   - Income section
'   - Error findings section
'   - Review findings (but not disposition code)
' ----------------------------------------------------------------------------
Private Sub ClearTANFSchedule(ByRef ws As Worksheet)
    On Error Resume Next
    
    ' Unprotect sheet to allow clearing
    ws.Unprotect Password:=SHEET_PASSWORD
    
    ' Clear person information (rows 30-44)
    ws.Range("A30:AP44").ClearContents
    
    ' Clear income section (rows 50-56)
    ws.Range("A50:AP56").ClearContents
    
    ' Clear error findings (rows 61-67)
    ws.Range("A61:AP67").ClearContents
    
    ' Clear review findings but keep disposition
    ws.Range("AL10").ClearContents
    ws.Range("AO10").ClearContents  ' Error amount
    
    ' Re-protect sheet
    ws.Protect Password:=SHEET_PASSWORD
    
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' Sub: ClearSNAPPosSchedule
' ----------------------------------------------------------------------------
' PURPOSE:
'   Clears data from a SNAP Positive schedule for a dropped review.
'
' PARAMETERS:
'   ws - The worksheet to clear
' ----------------------------------------------------------------------------
Private Sub ClearSNAPPosSchedule(ByRef ws As Worksheet)
    On Error Resume Next
    
    ws.Unprotect Password:=SHEET_PASSWORD
    
    ' Clear household composition section
    ws.Range("B89:AK122").ClearContents
    
    ' Clear income section
    ws.Range("B131:AK143").ClearContents
    
    ' Clear error section
    ws.Range("B149:AK155").ClearContents
    
    ' Clear findings but keep disposition
    ws.Range("K22").ClearContents
    
    ws.Protect Password:=SHEET_PASSWORD
    
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' Sub: ClearMAPosSchedule
' ----------------------------------------------------------------------------
' PURPOSE:
'   Clears data from an MA Positive schedule for a dropped review.
'
' PARAMETERS:
'   ws - The worksheet to clear
' ----------------------------------------------------------------------------
Private Sub ClearMAPosSchedule(ByRef ws As Worksheet)
    On Error Resume Next
    
    ws.Unprotect Password:=SHEET_PASSWORD
    
    ' Clear person-level information
    ws.Range("A51:AQ73").ClearContents
    
    ' Clear income section
    ws.Range("A78:AQ84").ClearContents
    
    ' Clear error findings
    ws.Range("A96:AQ112").ClearContents
    
    ' Clear review status but keep initial eligibility code
    ws.Range("S16").ClearContents
    
    ws.Protect Password:=SHEET_PASSWORD
    
    On Error GoTo 0
End Sub


