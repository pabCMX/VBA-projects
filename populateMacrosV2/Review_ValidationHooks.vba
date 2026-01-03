Attribute VB_Name = "Review_ValidationHooks"
' ============================================================================
' Review_ValidationHooks - Pre-Print Validation for Examiner Schedules
' ============================================================================
' WHAT THIS MODULE DOES:
'   This module contains the Workbook_BeforePrint event handler that validates
'   schedule data before the examiner prints. It catches common errors and
'   missing required fields BEFORE the schedule is submitted.
'
' WHY THIS EXISTS:
'   Examiners sometimes forget to fill in required fields. By validating at
'   print time (when they're about to submit), we catch errors early and
'   reduce the number of reviews returned for corrections.
'
' HOW IT WORKS:
'   When an examiner clicks Print (or uses Print Preview), this code runs
'   BEFORE the print dialog appears. If validation fails:
'   1. A message box explains what's wrong
'   2. Cancel = True prevents the print
'   3. The examiner can fix the issue and try again
'
' VALIDATION RULES BY PROGRAM:
'
'   SNAP POSITIVE:
'   - Homeless indicator (AJ76) required when disposition = 1
'   - Cert period length (T50) must be > 0 when disposition = 1
'   - Authorized rep (S55) required when disposition = 1
'   - SUA usage (W82) and utilities (AH82) must be consistent
'   - Allotment adjustment (AB50) required when disposition = 1
'   - Dependent care (section 4/5) sums must match deduction
'
'   MA POSITIVE:
'   - Citizenship/Identity code (O118) - commented out, was required
'   - Voter ID code (T118) - commented out
'   - IEVS code (Y118) - commented out
'   - Renewal type (AE118) - commented out
'
' IMPORTANT NOTES:
'   - This code is COPIED to each examiner schedule during population
'   - It runs in the examiner's schedule workbook, not the Populate workbook
'   - Keep it self-contained (no dependencies on other modules)
'   - Be careful with changes - affects all future schedules
'
' CHANGE LOG:
'   2026-01-03  Created from ThisWorkbook.vba
'               Added Option Explicit, V2 comments, modular structure
' ============================================================================

Option Explicit

' ============================================================================
' MAIN EVENT HANDLER
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: Workbook_BeforePrint
' ----------------------------------------------------------------------------
' PURPOSE:
'   Event handler that runs before any print operation.
'   Validates the active sheet if it's a schedule sheet.
'
' PARAMETERS:
'   Cancel - Set to True to prevent the print operation
' ----------------------------------------------------------------------------
Public Sub Workbook_BeforePrint(Cancel As Boolean)
    On Error Resume Next
    
    Dim wsSchedule As Worksheet
    Dim program As String
    Dim ws As Worksheet
    
    ' ========================================================================
    ' 1. FIND SCHEDULE SHEET AND DETERMINE PROGRAM
    ' ========================================================================
    
    ' Look for the schedule sheet (named with review number)
    For Each ws In ThisWorkbook.Worksheets
        ' SNAP Positive: 50xxxx, 51xxxx, 55xxxx
        If Left(ws.Name, 2) = "50" Or Left(ws.Name, 2) = "51" Or Left(ws.Name, 2) = "55" Then
            Set wsSchedule = ws
            program = "SNAP"
            Exit For
        ' MA Positive: 20xxxx, 21xxxx, 23xxxx
        ElseIf Left(ws.Name, 2) = "20" Or Left(ws.Name, 2) = "21" Or Left(ws.Name, 2) = "23" Then
            Set wsSchedule = ws
            program = "MA"
            Exit For
        End If
    Next ws
    
    ' If no schedule found or not printing the schedule, skip validation
    If wsSchedule Is Nothing Then Exit Sub
    If wsSchedule.Name <> ActiveSheet.Name Then Exit Sub
    
    ' ========================================================================
    ' 2. VALIDATE BASED ON PROGRAM
    ' ========================================================================
    
    Select Case program
        Case "SNAP"
            Cancel = Not ValidateSNAPPositive(wsSchedule)
        Case "MA"
            Cancel = Not ValidateMAPositive(wsSchedule)
    End Select
    
    On Error GoTo 0
End Sub


' ============================================================================
' SNAP POSITIVE VALIDATION
' ============================================================================

' ----------------------------------------------------------------------------
' Function: ValidateSNAPPositive
' ----------------------------------------------------------------------------
' PURPOSE:
'   Validates SNAP Positive schedule before printing.
'
' RETURNS:
'   True if validation passes, False if there are errors
' ----------------------------------------------------------------------------
Private Function ValidateSNAPPositive(ByRef ws As Worksheet) As Boolean
    On Error Resume Next
    
    Dim disposition As Variant
    Dim findingCode As Variant
    Dim i As Long
    Dim depCareSum As Double
    
    ValidateSNAPPositive = True  ' Assume valid until proven otherwise
    
    disposition = ws.Range("C22").Value
    findingCode = ws.Range("K22").Value
    
    ' ========================================================================
    ' VALIDATION: Homeless indicator required (disposition = 1)
    ' ========================================================================
    If disposition = "1" Or disposition = 1 Then
        If ws.Range("AJ76").Value = "" Then
            MsgBox "Please enter a value into item 42 (Homeless).", vbExclamation
            ValidateSNAPPositive = False
            Exit Function
        End If
    End If
    
    ' ========================================================================
    ' VALIDATION: Cert period length > 0 (disposition = 1)
    ' ========================================================================
    If disposition = "1" Or disposition = 1 Then
        Dim certPeriod As Variant
        certPeriod = ws.Range("T50").Value
        If certPeriod = "0" Or certPeriod = "" Or certPeriod = "00" Then
            MsgBox "Please enter a value greater than 0 into item 22 (length of cert. period).", vbExclamation
            ws.Range("T50").Value = ""
            ValidateSNAPPositive = False
            Exit Function
        End If
    End If
    
    ' ========================================================================
    ' VALIDATION: Authorized Representative required (disposition = 1)
    ' ========================================================================
    If disposition = "1" Or disposition = 1 Then
        If ws.Range("S55").Value = "" Then
            MsgBox "Please enter a value into item 27 (Authorized Representative).", vbExclamation
            ValidateSNAPPositive = False
            Exit Function
        End If
    End If
    
    ' ========================================================================
    ' VALIDATION: SUA consistency check
    ' ========================================================================
    ' If utilities (AH82) = 0, then SUA usage (W82) should be 1 and proration = "-"
    If disposition = "1" Or disposition = 1 Or disposition = "4" Or disposition = 4 Then
        If Val(ws.Range("AH82").Value) = 0 Then
            ws.Range("W82").Value = 1
            ws.Range("AA82").Value = "-"
        End If
    End If
    
    ' If SUA usage = 1 but utilities > 0, that's an error
    If disposition = "1" Or disposition = 1 Or disposition = "4" Or disposition = 4 Then
        If Val(ws.Range("AH82").Value) <> 0 And Val(ws.Range("W82").Value) = 1 Then
            MsgBox "Box 1 of Item 44 (Use of SUA) cannot = 1, if Item 45 (Utilities) is greater than 0", vbExclamation
            ws.Range("W82").Value = ""
            ws.Range("AA82").Value = ""
            ValidateSNAPPositive = False
            Exit Function
        ElseIf Val(ws.Range("W82").Value) <> 1 And _
               (ws.Range("AA82").Value = "-" Or ws.Range("AA82").Value = "") Then
            MsgBox "If Box 1 of Item 44 (Usage of SUA) is not = 1, then Box 2 of Item 44 (Proration of SUA) must be a 1 or 2.", vbExclamation
            ws.Range("AA82").Value = ""
            ValidateSNAPPositive = False
            Exit Function
        End If
    End If
    
    ' ========================================================================
    ' VALIDATION: Allotment adjustment required (disposition = 1)
    ' ========================================================================
    If disposition = "1" Or disposition = 1 Then
        If ws.Range("AB50").Value = "" Then
            MsgBox "Please enter a value into item 23 (Allotment Adjustment).", vbExclamation
            ValidateSNAPPositive = False
            Exit Function
        End If
    End If
    
    ' ========================================================================
    ' VALIDATION: Dependent care sum check (finding codes 1, 2, or 3)
    ' ========================================================================
    If findingCode = 1 Or findingCode = 2 Or findingCode = 3 Then
        depCareSum = 0
        For i = 89 To 122 Step 3
            If Val(ws.Range("E" & i).Value) = 1 Then
                depCareSum = depCareSum + Val(ws.Range("AJ" & i).Value)
            End If
        Next i
        
        If depCareSum < Val(ws.Range("O76").Value) - 5 Then
            MsgBox "Sum of Item 58 (Dependent Care Cost) which equals $" & depCareSum & _
                   " must be greater than or $5 less than Item 39 (Dependent Care Deduction) which equals $" & _
                   ws.Range("O76").Value & ".", vbExclamation
            ValidateSNAPPositive = False
            Exit Function
        End If
    ElseIf findingCode = "4" Or findingCode = 4 Then
        ' Ineligible review - prompt to clear sections 4 and 5
        If ws.Range("B89").Value <> "" Or ws.Range("B131").Value <> "" Then
            Dim result As VbMsgBoxResult
            result = MsgBox("This schedule is an ineligible review. Please click Yes to clear sections 4 and 5.", vbYesNo)
            If result = vbYes Then
                ws.Range("A89:AN122").Value = ""
                ws.Range("A131:AN143").Value = ""
                ValidateSNAPPositive = True
            Else
                ValidateSNAPPositive = False
                Exit Function
            End If
        End If
    End If
    
    On Error GoTo 0
End Function


' ============================================================================
' MA POSITIVE VALIDATION
' ============================================================================

' ----------------------------------------------------------------------------
' Function: ValidateMAPositive
' ----------------------------------------------------------------------------
' PURPOSE:
'   Validates MA Positive schedule before printing.
'   Note: Most validations are currently commented out in original code.
'
' RETURNS:
'   True if validation passes
' ----------------------------------------------------------------------------
Private Function ValidateMAPositive(ByRef ws As Worksheet) As Boolean
    On Error Resume Next
    
    ValidateMAPositive = True  ' Assume valid
    
    ' ========================================================================
    ' NOTE: The following validations were in the original code but are
    ' currently commented out. They can be enabled as needed.
    ' ========================================================================
    
    ' Citizenship and Identity code
    ' If ws.Range("F16").Value = "1" Then
    '     If ws.Range("O118").Value = "" Then
    '         MsgBox "Please enter a Citizenship and Identity code in the supplemental section."
    '         ValidateMAPositive = False
    '     End If
    ' End If
    
    ' Voter ID
    ' If ws.Range("F16").Value = "1" Then
    '     If ws.Range("T118").Value = "" Then
    '         MsgBox "Please enter a Voter ID code in the supplemental section."
    '         ValidateMAPositive = False
    '     End If
    ' End If
    
    ' IEVS type
    ' If ws.Range("D96").Value <> "" Then
    '     If ws.Range("Y118").Value = "" Then
    '         MsgBox "Please enter an IEVS code in the supplemental section."
    '         ValidateMAPositive = False
    '     End If
    ' End If
    
    ' Renewal type
    ' If ws.Range("F16").Value = "1" Then
    '     If ws.Range("AE118").Value = "" Then
    '         MsgBox "Please enter a Renewal Type code in the supplemental section."
    '         ValidateMAPositive = False
    '     End If
    ' End If
    
    On Error GoTo 0
End Function


