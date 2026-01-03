Attribute VB_Name = "Review_FormRouter"
' ============================================================================
' Review_FormRouter - Centralized Form Selection and Routing
' ============================================================================
' WHAT THIS MODULE DOES:
'   This module provides a centralized routing mechanism for form generation.
'   Instead of having complex Select Case logic scattered throughout UserForms,
'   all form routing goes through this module.
'
' WHY THIS EXISTS:
'   The original codebase had form routing logic duplicated in multiple places:
'   - SelectForms.frm
'   - MASelectForms.frm
'   - Various module code
'
'   By centralizing routing here:
'   - Single point of maintenance for adding/changing forms
'   - Consistent naming conventions
'   - Easier to trace which function handles which form
'   - UserForms just pass form names, don't need to know implementation
'
' HOW TO USE:
'   From a UserForm or other module:
'       Call GenerateForm("Findings Memo")
'       Call GenerateForm("CAO Appointment Letter")
'
'   The GenerateForm sub handles routing to the appropriate function.
'
' FORM NAME MAPPING:
'   Form names used in ListBoxes map to generation functions as follows:
'
'   SNAP Positive:
'     "Findings Memo"              -> GenerateFindingsMemo
'     "CAO Appointment Letter"     -> GenerateCAOAppointment
'     "Case Summary Template"      -> GenerateCaseSummary
'     "Telephone Appointment"      -> GenerateTelephoneAppointment
'     ... (see full list below)
'
' CHANGE LOG:
'   2026-01-03  Initial creation for V2 refactor
' ============================================================================

Option Explicit


' ============================================================================
' MAIN ROUTING FUNCTION
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: GenerateForm
' ----------------------------------------------------------------------------
' PURPOSE:
'   Main entry point for form generation. Routes the form name to the
'   appropriate generation function.
'
' PARAMETERS:
'   formName    - The display name of the form (as shown in ListBox)
'   memoType    - Optional: Additional parameter for forms that need subtypes
'                 (e.g., MA Finding Memo needs "MA_Pos_Find" or "MA_Neg_Find")
'
' EXAMPLE:
'   Call GenerateForm("Findings Memo")
'   Call GenerateForm("MA Findings Memo", "MA_Pos_Find")
' ----------------------------------------------------------------------------
Public Sub GenerateForm(ByVal formName As String, Optional ByVal memoType As String = "")
    On Error GoTo ErrorHandler
    
    Select Case formName
    
        ' ==================================================================
        ' FINDINGS MEMOS
        ' ==================================================================
        Case "Findings Memo"
            Call Finding_Memo_sub
            
        Case "MA Findings Memo"
            If memoType <> "" Then
                Call MA_Finding_Memo_sub(memoType)
            Else
                Call MA_Finding_Memo_sub("MA_Pos_Find")
            End If
            
        Case "Deficiency Memo"
            ' Route based on program type
            Select Case Left(ActiveSheet.Name, 1)
                Case "2"
                    Call MA_Finding_Memo_sub("MA_Pos_Def")
                Case "8"
                    Call MA_Finding_Memo_sub("MA_Neg_Def")
            End Select
            
        Case "Timeliness Information/Finding Memo", "Timeliness Memo"
            Call timeliness_switch
            
        Case "Potential Error Call Memo"
            Call PotentialErrorCall
            
        ' ==================================================================
        ' APPOINTMENT LETTERS
        ' ==================================================================
        Case "CAO Appointment Letter"
            Call CAOAppt
            
        Case "Spanish CAO Appointment Letter"
            Call SpCAOAppt
            
        Case "Telephone Appointment Letter"
            Call TeleAppt
            
        Case "Spanish Telephone Appointment Letter"
            Call SpTeleAppt
            
        Case "MA Appointment Letter"
            Call MAAppt
            
        Case "Spanish MA App. Letter", "Spanish MA Appointment Letter"
            Call MASpCAOAppt
            
        ' ==================================================================
        ' INFORMATION MEMOS
        ' ==================================================================
        Case "Information Memo"
            Call Info
            
        Case "Negative Info Memo"
            Call NewNeg
            
        Case "Post Office Memo"
            Call Post
            
        Case "Error Under Threshold Memo"
            Call Threshold
            
        ' ==================================================================
        ' PENDING LETTERS
        ' ==================================================================
        Case "Pending Letter"
            Call SNAPPend
            
        Case "Spanish Pending Letter"
            Call SpPend
            
        Case "TANF Pending"
            Call TANFPend
            
        Case "PA472-Pending Letter"
            Call PendLet
            
        ' ==================================================================
        ' QC14 SERIES
        ' ==================================================================
        Case "SNAP QC14F Memo"
            Call QC14F
            
        Case "SNAP QC14R Memo"
            Call QC14R
            
        Case "SNAP QC14C Memo", "QC14 Coop Memo"
            Call QC14C
            
        Case "QC14 CAO Request Memo"
            Call QC14
            
        Case "QC14 LTC Waiver Memo"
            Call MA_LTC_WAIVER
            
        Case "QC15 Memo"
            Call QC15
            
        ' ==================================================================
        ' CASH PROGRAM FORMS
        ' ==================================================================
        Case "Case Summary Template"
            Call Rush
            
        Case "Drop Worksheet"
            Call SNAPDrop
            
        Case "CAO Forms Request"
            Call SNAPCAORequest
            
        Case "TANF CAO Request Form"
            Call TANFCAORequest
            
        Case "AMR", "AMR Memo"
            Call AMR
            
        Case "Criminal History", "Criminal History Memo"
            Call Criminal
            
        Case "Notification Requirement Info Memo"
            Call TANF_Signature_Notification(1)
            
        Case "SAVE", "SAVE Deficiency Memo"
            Select Case Left(ActiveSheet.Name, 1)
                Case "2"
                    Call MA_Finding_Memo_sub("MA_Pos_SAVE")
                Case "1", "9"
                    Call TANF_SAVE
            End Select
            
        Case "School Verification"
            Call TANFSchool
            
        ' ==================================================================
        ' MA-SPECIFIC FORMS
        ' ==================================================================
        Case "Community Spouse Questionaire"
            Call ComSpouse
            
        Case "Preliminary Information Memo"
            Call Prelim_Info
            
        Case "NH LRR"
            Call NHLLR
            
        Case "NH Home Business"
            Call NHBUS
            
        Case "PA472-Self Emp."
            Call SelfEmp
            
        Case "Self Emp."
            Call SelfEmpDet
            
        Case "PA76"
            Call PA76
            
        Case "PA78", "Employment/Earnings Report"
            Select Case Left(ActiveSheet.Name, 1)
                Case "5"
                    Call SNAPPA78
                Case "1", "9"
                    Call TANFPA78
                Case Else
                    Call PA78
            End Select
            
        Case "PA83-Z"
            Call PA83Z
            
        Case "Household Composition"
            Call HouseholdComp
            
        Case "MA Support Form"
            Call MA_Supp
            
        Case "MA Waiver Memo"
            Call MA_WAIVER
            
        Case "Adoption Foster Care"
            Call Adoption
            
        Case "Zero Income Request"
            Call MA_Zero
            
        Case "Funeral Home Letter"
            Call Funeral
            
        ' ==================================================================
        ' LEP
        ' ==================================================================
        Case "LEP"
            Call LEP
            
        ' ==================================================================
        ' DEFAULT
        ' ==================================================================
        Case Else
            MsgBox "Form '" & formName & "' is not recognized.", _
                   vbExclamation, "Unknown Form"
    End Select
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error generating form '" & formName & "': " & Err.Description, _
           vbCritical, "Error"
    LogError "GenerateForm", Err.Number, Err.Description, formName
End Sub


' ============================================================================
' FORM LIST FUNCTIONS
' ============================================================================

' ----------------------------------------------------------------------------
' Function: GetFormList
' ----------------------------------------------------------------------------
' PURPOSE:
'   Returns an array of form names available for the specified program type.
'   Used to populate the ListBox in UF_SelectForms.
'
' PARAMETERS:
'   programType - The ProgramType enum value
'
' RETURNS:
'   Variant - Array of form name strings
'
' EXAMPLE:
'   Dim forms As Variant
'   forms = GetFormList(PROG_SNAP_POS)
'   For i = LBound(forms) To UBound(forms)
'       ListBox1.AddItem forms(i)
'   Next i
' ----------------------------------------------------------------------------
Public Function GetFormList(ByVal programType As ProgramType) As Variant
    
    Select Case programType
    
        Case PROG_SNAP_POS
            GetFormList = Array( _
                "Findings Memo", _
                "CAO Appointment Letter", _
                "Case Summary Template", _
                "Telephone Appointment Letter", _
                "SNAP QC14F Memo", _
                "SNAP QC14R Memo", _
                "SNAP QC14C Memo", _
                "Timeliness Information/Finding Memo", _
                "Information Memo", _
                "Post Office Memo", _
                "Pending Letter", _
                "Drop Worksheet", _
                "CAO Forms Request", _
                "Error Under Threshold Memo", _
                "LEP", _
                "Spanish CAO Appointment Letter", _
                "Spanish Telephone Appointment Letter", _
                "PA78", _
                "Spanish Pending Letter")
                
        Case PROG_SNAP_NEG
            GetFormList = Array( _
                "Findings Memo", _
                "Case Summary Template", _
                "Negative Info Memo", _
                "LEP")
                
        Case PROG_TANF
            GetFormList = Array( _
                "Findings Memo", _
                "Information Memo", _
                "Potential Error Call Memo", _
                "Notification Requirement Info Memo", _
                "AMR", _
                "Criminal History", _
                "TANF CAO Request Form", _
                "SAVE", _
                "Employment/Earnings Report", _
                "LEP", _
                "TANF Pending", _
                "School Verification")
                
        Case PROG_GA
            GetFormList = Array( _
                "Findings Memo", _
                "AMR Memo", _
                "Criminal History Memo", _
                "Information Memo", _
                "Potential Error Call Memo", _
                "SAVE", _
                "Employment/Earnings Report")
                
        Case PROG_MA_POS
            GetFormList = Array( _
                "Findings Memo", _
                "Deficiency Memo", _
                "MA Appointment Letter", _
                "Community Spouse Questionaire", _
                "Potential Error Call Memo", _
                "Information Memo", _
                "QC14 Coop Memo", _
                "QC14 CAO Request Memo", _
                "QC15 Memo", _
                "Preliminary Information Memo", _
                "NH LRR", _
                "NH Home Business", _
                "PA472-Pending Letter", _
                "PA472-Self Emp.", _
                "Self Emp.", _
                "PA76", _
                "PA78", _
                "PA83-Z", _
                "Household Composition", _
                "SAVE Deficiency Memo", _
                "MA Support Form", _
                "MA Waiver Memo", _
                "Adoption Foster Care", _
                "Zero Income Request", _
                "QC14 LTC Waiver Memo", _
                "Funeral Home Letter", _
                "LEP", _
                "Spanish MA App. Letter")
                
        Case PROG_MA_NEG
            GetFormList = Array( _
                "Findings Memo", _
                "Deficiency Memo", _
                "Potential Error Call Memo", _
                "Information Memo", _
                "LEP")
                
        Case PROG_MA_PE
            GetFormList = Array( _
                "Findings Memo", _
                "Information Memo")
                
        Case Else
            GetFormList = Array()
    End Select
    
End Function

' ----------------------------------------------------------------------------
' Function: GetFormCount
' ----------------------------------------------------------------------------
' PURPOSE:
'   Returns the number of forms available for a program type.
'
' PARAMETERS:
'   programType - The ProgramType enum value
'
' RETURNS:
'   Long - Number of available forms
' ----------------------------------------------------------------------------
Public Function GetFormCount(ByVal programType As ProgramType) As Long
    Dim forms As Variant
    forms = GetFormList(programType)
    
    On Error Resume Next
    GetFormCount = UBound(forms) - LBound(forms) + 1
    If Err.Number <> 0 Then GetFormCount = 0
    On Error GoTo 0
End Function


' ============================================================================
' SHOW FORM DIALOGS
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: ShowSelectForms
' ----------------------------------------------------------------------------
' PURPOSE:
'   Displays the appropriate form selection dialog based on program type.
'   This is the main entry point called by buttons on schedules.
' ----------------------------------------------------------------------------
Public Sub ShowSelectForms()
    On Error Resume Next
    UF_SelectForms.Show
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' Sub: ShowMASelectForms
' ----------------------------------------------------------------------------
' PURPOSE:
'   Displays the MA-specific form selection dialog.
' ----------------------------------------------------------------------------
Public Sub ShowMASelectForms()
    On Error Resume Next
    UF_MA_SelectForms.Show
    On Error GoTo 0
End Sub



