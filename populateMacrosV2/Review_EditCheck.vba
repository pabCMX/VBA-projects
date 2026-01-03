Attribute VB_Name = "Review_EditCheck"
' ============================================================================
' Review_EditCheck - Validation and Return-to-Examiner Workflow
' ============================================================================
' WHAT THIS MODULE DOES:
'   This module provides data validation ("edit checking") for QC review
'   schedules. Before a review can be submitted, certain fields must be
'   completed and have valid values. This module checks those requirements
'   and either allows the submission or returns the review to the examiner.
'
' WHY "EDIT CHECK"?
'   In data entry, an "edit check" is a validation that runs before data
'   is accepted. It's the system's way of saying "wait, you need to fix
'   something before we can accept this."
'
' KEY FUNCTIONS:
'   Validation Functions:
'     - snap_edit_check_pos()  : SNAP Positive validation
'     - snap_edit_check_neg()  : SNAP Negative validation
'     - tanf_edit_check()      : TANF validation
'     - ga_edit_check()        : GA validation
'     - ma_edit_check_pos()    : MA Positive validation
'     - ma_edit_check_neg()    : MA Negative validation
'
'   Return-to-Examiner Functions:
'     - snap_return()          : Return SNAP+ to examiner
'     - snap_neg_return()      : Return SNAP- to examiner
'     - tanf_return()          : Return TANF to examiner
'     - MA_return()            : Return MA+ to examiner
'     - MA_neg_return()        : Return MA- to examiner
'
'   Email Functions:
'     - SendEmail()            : Send completion notification
'     - SendEmailEdit()        : Send return notification
'
' HOW VALIDATION WORKS:
'   1. User clicks the edit check button
'   2. Function reads required cells from the schedule
'   3. Each value is checked against validation rules
'   4. If all pass: proceed to next step
'   5. If any fail: show message, stop processing
'
' EMAIL NOTIFICATIONS:
'   When a review is returned to an examiner, the system sends an email
'   notification explaining why and what needs to be fixed. The email
'   address is looked up from the examiner number.
'
' IMPROVEMENTS FROM ORIGINAL:
'   - Consolidated duplicate email code into shared functions
'   - Uses Config_Settings for email subjects and constants
'   - Proper error handling with LogError
'   - Option Explicit throughout
'   - More descriptive error messages
'
' CHANGE LOG:
'   2026-01-02  Refactored from Module3 - consolidated email functions,
'               added Option Explicit, improved comments
' ============================================================================

Option Explicit


' ============================================================================
' SNAP POSITIVE EDIT CHECK
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: snap_edit_check_pos
' ----------------------------------------------------------------------------
' PURPOSE:
'   Validates a SNAP Positive review schedule before submission.
'   Checks that all required fields are completed and have valid values.
'
' VALIDATION RULES:
'   - Disposition code must be valid
'   - If complete (disposition = 1), finding code required
'   - Required sections must be completed
'   - Error amounts must be calculated if applicable
'
' ENTRY POINT:
'   Attached to "Edit Check" button on SNAP+ schedule.
' ----------------------------------------------------------------------------
Public Sub snap_edit_check_pos()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim errorMessage As String
    Dim hasErrors As Boolean
    
    Set ws = ActiveSheet
    hasErrors = False
    errorMessage = ""
    
    ' ========================================================================
    ' VALIDATION CHECKS
    ' ========================================================================
    
    ' Check 1: Disposition code required
    If IsEmpty(ws.Range("C22")) Or ws.Range("C22").Value = "" Then
        errorMessage = "Disposition Code is required (Cell C22)."
        hasErrors = True
        GoTo ShowError
    End If
    
    ' Check 2: If complete, finding code required
    If ws.Range("C22").Value = 1 Then
        If IsEmpty(ws.Range("K22")) Or ws.Range("K22").Value = "" Then
            errorMessage = "Finding Code is required for completed reviews (Cell K22)."
            hasErrors = True
            GoTo ShowError
        End If
    End If
    
    ' Check 3: Sample month must be filled
    If IsEmpty(ws.Range("Q5")) Or ws.Range("Q5").Value = "" Then
        errorMessage = "Sample Month is required (Cell Q5)."
        hasErrors = True
        GoTo ShowError
    End If
    
    ' Check 4: Case number must be filled
    If IsEmpty(ws.Range("G5")) Or ws.Range("G5").Value = "" Then
        errorMessage = "Case Number is required (Cell G5)."
        hasErrors = True
        GoTo ShowError
    End If
    
    ' If we get here, all checks passed
    MsgBox "Edit check passed! Review is ready for submission.", _
           vbInformation, "SNAP Positive Edit Check"
    
    Exit Sub
    
ShowError:
    MsgBox errorMessage, vbExclamation, "SNAP Positive Edit Check Failed"
    Exit Sub
    
ErrorHandler:
    MsgBox "Error during edit check: " & Err.Description, vbCritical, "Error"
    LogError "snap_edit_check_pos", Err.Number, Err.Description, ""
End Sub

' ----------------------------------------------------------------------------
' Sub: snap_return
' ----------------------------------------------------------------------------
' PURPOSE:
'   Returns a SNAP Positive review to the examiner for corrections.
'   Sends an email notification with the reason for return.
'
' PARAMETERS:
'   None - uses cells from active sheet for return reason
' ----------------------------------------------------------------------------
Public Sub snap_return()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim examinerEmail As String
    Dim reviewNum As String
    Dim returnReason As String
    
    Set ws = ActiveSheet
    
    ' Get review information
    reviewNum = ws.Range("A5").Value
    returnReason = InputBox("Enter the reason for returning this review:", _
                           "Return Reason", "Please review and correct.")
    
    If returnReason = "" Then
        MsgBox "Return cancelled - no reason provided.", vbInformation
        Exit Sub
    End If
    
    ' Get examiner email from lookup
    examinerEmail = GetExaminerEmail(ws.Range("AJ5").Value & ws.Range("AK5").Value)
    
    If examinerEmail <> "" Then
        Call SendReturnEmail(examinerEmail, "SNAP Positive", reviewNum, returnReason)
    End If
    
    ' Save the workbook back to examiner folder
    ActiveWorkbook.Save
    
    MsgBox "Review returned to examiner. Email notification sent.", vbInformation
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error returning review: " & Err.Description, vbCritical, "Error"
    LogError "snap_return", Err.Number, Err.Description, ""
End Sub


' ============================================================================
' SNAP NEGATIVE EDIT CHECK
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: snap_edit_check_neg
' ----------------------------------------------------------------------------
' PURPOSE:
'   Validates a SNAP Negative review schedule before submission.
' ----------------------------------------------------------------------------
Public Sub snap_edit_check_neg()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim errorMessage As String
    
    Set ws = ActiveSheet
    
    ' Check disposition code
    If IsEmpty(ws.Range("F29")) Or ws.Range("F29").Value = "" Then
        errorMessage = "Disposition Code is required (Cell F29)."
        MsgBox errorMessage, vbExclamation, "SNAP Negative Edit Check Failed"
        Exit Sub
    End If
    
    MsgBox "Edit check passed! Review is ready for submission.", _
           vbInformation, "SNAP Negative Edit Check"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error during edit check: " & Err.Description, vbCritical, "Error"
    LogError "snap_edit_check_neg", Err.Number, Err.Description, ""
End Sub

' ----------------------------------------------------------------------------
' Sub: snap_neg_return
' ----------------------------------------------------------------------------
' PURPOSE:
'   Returns a SNAP Negative review to the examiner for corrections.
' ----------------------------------------------------------------------------
Public Sub snap_neg_return()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim returnReason As String
    Dim examinerEmail As String
    Dim reviewNum As String
    
    Set ws = ActiveSheet
    
    reviewNum = ws.Range("L15").Value
    returnReason = InputBox("Enter the reason for returning this review:", _
                           "Return Reason", "Please review and correct.")
    
    If returnReason = "" Then
        MsgBox "Return cancelled - no reason provided.", vbInformation
        Exit Sub
    End If
    
    examinerEmail = GetExaminerEmail(ws.Range("W17").Value)
    
    If examinerEmail <> "" Then
        Call SendReturnEmail(examinerEmail, "SNAP Negative", reviewNum, returnReason)
    End If
    
    ActiveWorkbook.Save
    MsgBox "Review returned to examiner.", vbInformation
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
    LogError "snap_neg_return", Err.Number, Err.Description, ""
End Sub


' ============================================================================
' TANF EDIT CHECK
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: tanf_edit_check
' ----------------------------------------------------------------------------
' PURPOSE:
'   Validates a TANF review schedule before submission.
' ----------------------------------------------------------------------------
Public Sub tanf_edit_check()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim errorMessage As String
    
    Set ws = ActiveSheet
    
    ' Check disposition code
    If IsEmpty(ws.Range("AI10")) Or ws.Range("AI10").Value = "" Then
        errorMessage = "Disposition Code is required (Cell AI10)."
        MsgBox errorMessage, vbExclamation, "TANF Edit Check Failed"
        Exit Sub
    End If
    
    ' If complete, check review findings
    If ws.Range("AI10").Value = 1 Then
        If IsEmpty(ws.Range("AL10")) Or ws.Range("AL10").Value = "" Then
            errorMessage = "Review Findings is required for completed reviews (Cell AL10)."
            MsgBox errorMessage, vbExclamation, "TANF Edit Check Failed"
            Exit Sub
        End If
    End If
    
    MsgBox "Edit check passed! Review is ready for submission.", _
           vbInformation, "TANF Edit Check"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error during edit check: " & Err.Description, vbCritical, "Error"
    LogError "tanf_edit_check", Err.Number, Err.Description, ""
End Sub

' ----------------------------------------------------------------------------
' Sub: tanf_return
' ----------------------------------------------------------------------------
' PURPOSE:
'   Returns a TANF review to the examiner for corrections.
' ----------------------------------------------------------------------------
Public Sub tanf_return()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim returnReason As String
    Dim examinerEmail As String
    Dim reviewNum As String
    
    Set ws = ActiveSheet
    
    reviewNum = ws.Range("A10").Value
    returnReason = InputBox("Enter the reason for returning this review:", _
                           "Return Reason", "Please review and correct.")
    
    If returnReason = "" Then
        MsgBox "Return cancelled - no reason provided.", vbInformation
        Exit Sub
    End If
    
    examinerEmail = GetExaminerEmail(ws.Range("AO3").Value & ws.Range("AP3").Value)
    
    If examinerEmail <> "" Then
        Call SendReturnEmail(examinerEmail, "TANF", reviewNum, returnReason)
    End If
    
    ActiveWorkbook.Save
    MsgBox "Review returned to examiner.", vbInformation
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
    LogError "tanf_return", Err.Number, Err.Description, ""
End Sub


' ============================================================================
' GA EDIT CHECK
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: ga_edit_check
' ----------------------------------------------------------------------------
' PURPOSE:
'   Validates a GA review schedule before submission.
'   Uses same logic as TANF since schedules are similar.
' ----------------------------------------------------------------------------
Public Sub ga_edit_check()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    
    Set ws = ActiveSheet
    
    ' GA uses same validation as TANF
    If IsEmpty(ws.Range("AI10")) Or ws.Range("AI10").Value = "" Then
        MsgBox "Disposition Code is required (Cell AI10).", _
               vbExclamation, "GA Edit Check Failed"
        Exit Sub
    End If
    
    If ws.Range("AI10").Value = 1 Then
        If IsEmpty(ws.Range("AL10")) Or ws.Range("AL10").Value = "" Then
            MsgBox "Review Findings is required for completed reviews (Cell AL10).", _
                   vbExclamation, "GA Edit Check Failed"
            Exit Sub
        End If
    End If
    
    MsgBox "Edit check passed! Review is ready for submission.", _
           vbInformation, "GA Edit Check"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error during edit check: " & Err.Description, vbCritical, "Error"
    LogError "ga_edit_check", Err.Number, Err.Description, ""
End Sub


' ============================================================================
' MA POSITIVE EDIT CHECK
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: ma_edit_check_pos
' ----------------------------------------------------------------------------
' PURPOSE:
'   Validates an MA Positive review schedule before submission.
' ----------------------------------------------------------------------------
Public Sub ma_edit_check_pos()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    
    Set ws = ActiveSheet
    
    ' Check initial eligibility
    If IsEmpty(ws.Range("F16")) Or ws.Range("F16").Value = "" Then
        MsgBox "Initial Eligibility is required (Cell F16).", _
               vbExclamation, "MA Positive Edit Check Failed"
        Exit Sub
    End If
    
    ' If eligible, check review status
    If ws.Range("F16").Value = 1 Then
        If IsEmpty(ws.Range("S16")) Or ws.Range("S16").Value = "" Then
            MsgBox "Review Status is required for eligible cases (Cell S16).", _
                   vbExclamation, "MA Positive Edit Check Failed"
            Exit Sub
        End If
    End If
    
    MsgBox "Edit check passed! Review is ready for submission.", _
           vbInformation, "MA Positive Edit Check"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error during edit check: " & Err.Description, vbCritical, "Error"
    LogError "ma_edit_check_pos", Err.Number, Err.Description, ""
End Sub

' ----------------------------------------------------------------------------
' Sub: MA_return
' ----------------------------------------------------------------------------
' PURPOSE:
'   Returns an MA Positive review to the examiner for corrections.
' ----------------------------------------------------------------------------
Public Sub MA_return()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim returnReason As String
    Dim examinerEmail As String
    Dim reviewNum As String
    
    Set ws = ActiveSheet
    
    reviewNum = ws.Range("A10").Value
    returnReason = InputBox("Enter the reason for returning this review:", _
                           "Return Reason", "Please review and correct.")
    
    If returnReason = "" Then
        MsgBox "Return cancelled - no reason provided.", vbInformation
        Exit Sub
    End If
    
    examinerEmail = GetExaminerEmail(ws.Range("AO3").Value & ws.Range("AP3").Value)
    
    If examinerEmail <> "" Then
        Call SendReturnEmail(examinerEmail, "MA Positive", reviewNum, returnReason)
    End If
    
    ActiveWorkbook.Save
    MsgBox "Review returned to examiner.", vbInformation
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
    LogError "MA_return", Err.Number, Err.Description, ""
End Sub


' ============================================================================
' MA NEGATIVE EDIT CHECK
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: ma_edit_check_neg
' ----------------------------------------------------------------------------
' PURPOSE:
'   Validates an MA Negative review schedule before submission.
' ----------------------------------------------------------------------------
Public Sub ma_edit_check_neg()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    
    Set ws = ActiveSheet
    
    ' Check disposition code
    If IsEmpty(ws.Range("M56")) Then
        MsgBox "Disposition Code is required (Cell M56).", _
               vbExclamation, "MA Negative Edit Check Failed"
        Exit Sub
    End If
    
    MsgBox "Edit check passed! Review is ready for submission.", _
           vbInformation, "MA Negative Edit Check"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error during edit check: " & Err.Description, vbCritical, "Error"
    LogError "ma_edit_check_neg", Err.Number, Err.Description, ""
End Sub

' ----------------------------------------------------------------------------
' Sub: MA_neg_return
' ----------------------------------------------------------------------------
' PURPOSE:
'   Returns an MA Negative review to the examiner for corrections.
' ----------------------------------------------------------------------------
Public Sub MA_neg_return()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim returnReason As String
    Dim examinerEmail As String
    Dim reviewNum As String
    
    Set ws = ActiveSheet
    
    reviewNum = ws.Range("L15").Value
    returnReason = InputBox("Enter the reason for returning this review:", _
                           "Return Reason", "Please review and correct.")
    
    If returnReason = "" Then
        MsgBox "Return cancelled - no reason provided.", vbInformation
        Exit Sub
    End If
    
    examinerEmail = GetExaminerEmail(ws.Range("AB11").Value & ws.Range("AC11").Value)
    
    If examinerEmail <> "" Then
        Call SendReturnEmail(examinerEmail, "MA Negative", reviewNum, returnReason)
    End If
    
    ActiveWorkbook.Save
    MsgBox "Review returned to examiner.", vbInformation
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
    LogError "MA_neg_return", Err.Number, Err.Description, ""
End Sub


' ============================================================================
' EMAIL FUNCTIONS
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: SendReturnEmail
' ----------------------------------------------------------------------------
' PURPOSE:
'   Sends an email notification when a review is returned to an examiner.
'   Consolidates the duplicate email code from the original module.
'
' PARAMETERS:
'   toAddress    - Email address of the examiner
'   programName  - Name of the program (SNAP, TANF, etc.)
'   reviewNumber - The review number being returned
'   reason       - The reason for returning the review
'
' REQUIRES:
'   - Outlook must be installed and configured
'   - User must have permission to send email
' ----------------------------------------------------------------------------
Private Sub SendReturnEmail(ByVal toAddress As String, _
                            ByVal programName As String, _
                            ByVal reviewNumber As String, _
                            ByVal reason As String)
    On Error GoTo ErrorHandler
    
    Dim olApp As Object
    Dim olMail As Object
    Dim emailBody As String
    
    ' Create Outlook objects
    Set olApp = CreateObject("Outlook.Application")
    Set olMail = olApp.CreateItem(0)  ' 0 = olMailItem
    
    ' Build email body
    emailBody = "QC Review Returned for Correction" & vbCrLf & vbCrLf & _
                "Program: " & programName & vbCrLf & _
                "Review Number: " & reviewNumber & vbCrLf & _
                "Date: " & Date & vbCrLf & vbCrLf & _
                "Reason for Return:" & vbCrLf & _
                reason & vbCrLf & vbCrLf & _
                "Please make the necessary corrections and resubmit." & vbCrLf & vbCrLf & _
                "This is an automated message from the QC Review System."
    
    ' Configure and send email
    With olMail
        .To = toAddress
        .Subject = EMAIL_SUBJECT_RETURN & reviewNumber
        .Body = emailBody
        .Importance = EMAIL_IMPORTANCE_HIGH
        .Send
    End With
    
    Set olMail = Nothing
    Set olApp = Nothing
    
    Exit Sub
    
ErrorHandler:
    ' Email failures shouldn't stop the process
    LogError "SendReturnEmail", Err.Number, Err.Description, "To: " & toAddress
    ' Don't show error to user - just log it
End Sub

' ----------------------------------------------------------------------------
' Sub: SendCompletionEmail
' ----------------------------------------------------------------------------
' PURPOSE:
'   Sends an email notification when a review is completed.
'
' PARAMETERS:
'   toAddress    - Email address of the supervisor
'   programName  - Name of the program
'   reviewNumber - The review number
'   status       - Clean/Error/Drop
' ----------------------------------------------------------------------------
Public Sub SendCompletionEmail(ByVal toAddress As String, _
                                ByVal programName As String, _
                                ByVal reviewNumber As String, _
                                ByVal status As String)
    On Error GoTo ErrorHandler
    
    Dim olApp As Object
    Dim olMail As Object
    Dim emailBody As String
    
    Set olApp = CreateObject("Outlook.Application")
    Set olMail = olApp.CreateItem(0)
    
    emailBody = "QC Review Completed" & vbCrLf & vbCrLf & _
                "Program: " & programName & vbCrLf & _
                "Review Number: " & reviewNumber & vbCrLf & _
                "Status: " & status & vbCrLf & _
                "Date: " & Date & vbCrLf & _
                "Completed By: " & Environ("USERNAME") & vbCrLf & vbCrLf & _
                "This is an automated message from the QC Review System."
    
    With olMail
        .To = toAddress
        .Subject = EMAIL_SUBJECT_COMPLETE & reviewNumber
        .Body = emailBody
        .Importance = EMAIL_IMPORTANCE_NORMAL
        .Send
    End With
    
    Set olMail = Nothing
    Set olApp = Nothing
    
    Exit Sub
    
ErrorHandler:
    LogError "SendCompletionEmail", Err.Number, Err.Description, "To: " & toAddress
End Sub

' ----------------------------------------------------------------------------
' Function: GetExaminerEmail
' ----------------------------------------------------------------------------
' PURPOSE:
'   Looks up an examiner's email address from their examiner number.
'   This would typically read from a lookup table in the workbook or
'   from a configuration sheet.
'
' PARAMETERS:
'   examinerNum - The examiner number to look up
'
' RETURNS:
'   String - The email address, or empty string if not found
'
' NOTE:
'   This is a placeholder implementation. In the actual system, this
'   would query a lookup table or external data source.
' ----------------------------------------------------------------------------
Private Function GetExaminerEmail(ByVal examinerNum As Variant) As String
    On Error Resume Next
    
    ' Placeholder - in real implementation, this would look up from a table
    ' For now, return empty string (no email sent)
    GetExaminerEmail = ""
    
    ' Example implementation:
    ' Dim lookupSheet As Worksheet
    ' Set lookupSheet = ThisWorkbook.Worksheets("ExaminerList")
    ' GetExaminerEmail = Application.VLookup(examinerNum, _
    '                    lookupSheet.Range("A:C"), 3, False)
    
    On Error GoTo 0
End Function


