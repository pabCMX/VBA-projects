Attribute VB_Name = "Review_CashMemos"
' ============================================================================
' Review_CashMemos - Cash Assistance Memo Document Generation
' ============================================================================
' WHAT THIS MODULE DOES:
'   Generates various memos and letters related to cash assistance programs
'   (TANF, GA, SNAP, MA). These documents communicate with clients and CAOs
'   about case status, required actions, and review findings.
'
' MIGRATED FROM V1:
'   CashMemos.vba - Subs: AMR(), Criminal(), Info(), TANF_Signature_Notification(),
'                   MA_Zero(), PendLet(), SelfEmp(), SelfEmpDet(), PA76(), PA78(),
'                   SNAPPA78(), TANFPA78(), MA_Supp(), PA83Z(), HouseholdComp(),
'                   Threshold(), MA_SAVE(), Funeral(), Adoption(), TANF_SAVE(),
'                   LEP(), TANFPend(), TANFSchool()
'
' V1 -> V2 FUNCTION MAPPING:
'   AMR()                      -> GenerateAMR()
'   Criminal()                 -> GenerateCriminalHistory()
'   Info()                     -> GenerateInformationMemo()
'   TANF_Signature_Notification() -> GenerateTANFSignatureNotification()
'   MA_Zero()                  -> GenerateMAZeroIncome()
'   PendLet()                  -> GeneratePendingLetter()
'   SelfEmp()                  -> GenerateSelfEmploymentBasic()
'   SelfEmpDet()               -> GenerateSelfEmploymentDetailed()
'   PA76()                     -> GenerateMASupportForm()
'   PA78()                     -> GenerateHouseholdComposition()
'   SNAPPA78()                 -> GenerateSNAPHouseholdComp()
'   TANFPA78()                 -> GenerateTANFHouseholdComp()
'   MA_Supp()                  -> GenerateMASupplementForm()
'   PA83Z()                    -> GenerateThresholdMemo()
'   HouseholdComp()            -> GenerateGeneralHouseholdComp()
'   Threshold()                -> GenerateErrorThresholdMemo()
'   MA_SAVE()                  -> GenerateMASAVE()
'   Funeral()                  -> GenerateFuneralHomeLetter()
'   Adoption()                 -> GenerateAdoptionFosterCare()
'   TANF_SAVE()                -> GenerateTANFSAVE()
'   LEP()                      -> GenerateLEP()
'   TANFPend()                 -> GenerateTANFPending()
'   TANFSchool()               -> GenerateSchoolVerification()
'
' CHANGE LOG:
'   2026-01-03  Refactored from CashMemos.vba - added Option Explicit,
'               centralized network detection, V2 commenting style
' ============================================================================

Option Explicit


' ============================================================================
' CORE MEMO GENERATION FUNCTIONS
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: GenerateAMR
' ----------------------------------------------------------------------------
' PURPOSE:
'   Generates an AMR (Administrative Monthly Report) memo for TANF/GA reviews.
'   Uses mail merge to populate a Word template with case data.
'
' V1 EQUIVALENT: AMR()
' ----------------------------------------------------------------------------
Public Sub GenerateAMR()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim programType As ProgramType
    Dim dqcPath As String
    Dim templatePath As String
    Dim outputPath As String
    Dim reviewNumber As String
    Dim sampleMonth As String
    
    Set ws = ActiveSheet
    programType = GetProgramFromSheetName(ws.Name)
    
    ' Validate program type
    If programType <> PROG_TANF And programType <> PROG_GA Then
        MsgBox "AMR memo is only available for TANF and GA reviews.", _
               vbInformation, "Not Applicable"
        Exit Sub
    End If
    
    ' Get network path
    dqcPath = GetDQCDriveLetterOrError()
    
    ' Build template path
    templatePath = dqcPath & "Finding Memo\AMR memo.docx"
    
    ' Validate template exists
    If Not PathExists(templatePath) Then
        MsgBox "AMR template not found at: " & templatePath, _
               vbCritical, "Template Not Found"
        Exit Sub
    End If
    
    ' Get review details
    reviewNumber = CStr(ws.Range("A10").Value)
    sampleMonth = CStr(ws.Range("AB10").Value)
    
    ' Build output filename
    outputPath = ActiveWorkbook.Path & "\AMR for Review Number " & _
                 reviewNumber & " for Sample Month " & sampleMonth & ".doc"
    
    ' Check if file already exists
    If Dir(outputPath) <> "" Then
        If MsgBox("An AMR memo already exists. Overwrite?", vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If
    
    ' Generate the memo
    Call CreateMemoFromTemplate(ws, programType, templatePath, outputPath, "AMR")
    
    MsgBox "AMR Memo has been saved to: " & outputPath, vbInformation, "Success"
    
    Exit Sub
    
ErrorHandler:
    LogError "GenerateAMR", Err.Number, Err.Description, ""
    MsgBox "Error generating AMR memo: " & Err.Description, vbCritical, "Error"
End Sub

' ----------------------------------------------------------------------------
' Sub: GenerateCriminalHistory
' ----------------------------------------------------------------------------
' PURPOSE:
'   Generates a Criminal History memo for cases involving criminal history checks.
'
' V1 EQUIVALENT: Criminal()
' ----------------------------------------------------------------------------
Public Sub GenerateCriminalHistory()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim programType As ProgramType
    Dim dqcPath As String
    Dim templatePath As String
    Dim outputPath As String
    Dim reviewNumber As String
    Dim sampleMonth As String
    
    Set ws = ActiveSheet
    programType = GetProgramFromSheetName(ws.Name)
    
    ' Get network path
    dqcPath = GetDQCDriveLetterOrError()
    
    ' Build template path
    templatePath = dqcPath & "Finding Memo\Criminal History.docx"
    
    ' Validate template exists
    If Not PathExists(templatePath) Then
        MsgBox "Criminal History template not found at: " & templatePath, _
               vbCritical, "Template Not Found"
        Exit Sub
    End If
    
    ' Get review details based on program type
    Select Case programType
        Case PROG_TANF, PROG_GA
            reviewNumber = CStr(ws.Range("A10").Value)
            sampleMonth = CStr(ws.Range("AB10").Value)
        Case PROG_SNAP_POS
            reviewNumber = CStr(ws.Range("A18").Value)
            sampleMonth = CStr(ws.Range("AD18").Value) & CStr(ws.Range("AG18").Value)
        Case Else
            reviewNumber = ws.Name
            sampleMonth = Format(Date, "MMYYYY")
    End Select
    
    ' Build output filename
    outputPath = ActiveWorkbook.Path & "\Criminal History Memo for Review Number " & _
                 reviewNumber & " for Sample Month " & sampleMonth & ".doc"
    
    ' Check if file already exists
    If Dir(outputPath) <> "" Then
        If MsgBox("A Criminal History memo already exists. Overwrite?", vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If
    
    ' Generate the memo
    Call CreateMemoFromTemplate(ws, programType, templatePath, outputPath, "Criminal History")
    
    MsgBox "Criminal History Memo has been saved to: " & outputPath, vbInformation, "Success"
    
    Exit Sub
    
ErrorHandler:
    LogError "GenerateCriminalHistory", Err.Number, Err.Description, ""
    MsgBox "Error generating Criminal History memo: " & Err.Description, vbCritical, "Error"
End Sub

' ----------------------------------------------------------------------------
' Sub: GenerateInformationMemo
' ----------------------------------------------------------------------------
' PURPOSE:
'   Generates a general Information memo for all program types.
'   This is used to request additional information from CAOs.
'
' V1 EQUIVALENT: Info()
' ----------------------------------------------------------------------------
Public Sub GenerateInformationMemo()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim programType As ProgramType
    Dim dqcPath As String
    Dim templatePath As String
    Dim outputPath As String
    Dim reviewNumber As String
    Dim sampleMonth As String
    
    Set ws = ActiveSheet
    programType = GetProgramFromSheetName(ws.Name)
    
    ' Get network path
    dqcPath = GetDQCDriveLetterOrError()
    
    ' Build template path based on program
    Select Case programType
        Case PROG_TANF
            templatePath = dqcPath & "Finding Memo\TANF Info Memo.docx"
        Case PROG_GA
            templatePath = dqcPath & "Finding Memo\GA Info Memo.docx"
        Case PROG_SNAP_POS
            templatePath = dqcPath & "Finding Memo\SNAP Pos Info Memo.docx"
        Case PROG_SNAP_NEG
            templatePath = dqcPath & "Finding Memo\SNAP Neg Info Memo.docx"
        Case PROG_MA_POS
            templatePath = dqcPath & "Finding Memo\MA Pos Info Memo.docx"
        Case PROG_MA_NEG
            templatePath = dqcPath & "Finding Memo\MA Neg Info Memo.docx"
        Case Else
            templatePath = dqcPath & "Finding Memo\General Info Memo.docx"
    End Select
    
    ' Validate template exists
    If Not PathExists(templatePath) Then
        MsgBox "Information memo template not found at: " & templatePath, _
               vbCritical, "Template Not Found"
        Exit Sub
    End If
    
    ' Get review details
    reviewNumber = GetReviewNumber(ws, programType)
    sampleMonth = GetSampleMonth(ws, programType)
    
    ' Build output filename
    outputPath = ActiveWorkbook.Path & "\Information Memo for Review Number " & _
                 reviewNumber & " for Sample Month " & sampleMonth & ".doc"
    
    ' Check if file already exists
    If Dir(outputPath) <> "" Then
        If MsgBox("An Information memo already exists. Overwrite?", vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If
    
    ' Generate the memo
    Call CreateMemoFromTemplate(ws, programType, templatePath, outputPath, "Information")
    
    MsgBox "Information Memo has been saved to: " & outputPath, vbInformation, "Success"
    
    Exit Sub
    
ErrorHandler:
    LogError "GenerateInformationMemo", Err.Number, Err.Description, ""
    MsgBox "Error generating Information memo: " & Err.Description, vbCritical, "Error"
End Sub

' ----------------------------------------------------------------------------
' Sub: GenerateTANFSignatureNotification
' ----------------------------------------------------------------------------
' PURPOSE:
'   Generates TANF Signature Notification letter.
'
' V1 EQUIVALENT: TANF_Signature_Notification()
' ----------------------------------------------------------------------------
Public Sub GenerateTANFSignatureNotification()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim programType As ProgramType
    Dim dqcPath As String
    Dim templatePath As String
    Dim outputPath As String
    
    Set ws = ActiveSheet
    programType = GetProgramFromSheetName(ws.Name)
    
    If programType <> PROG_TANF Then
        MsgBox "TANF Signature Notification is only available for TANF reviews.", _
               vbInformation, "Not Applicable"
        Exit Sub
    End If
    
    dqcPath = GetDQCDriveLetterOrError()
    templatePath = dqcPath & "Finding Memo\TANF Signature Notification.docx"
    
    If Not PathExists(templatePath) Then
        MsgBox "Template not found at: " & templatePath, vbCritical, "Template Not Found"
        Exit Sub
    End If
    
    outputPath = ActiveWorkbook.Path & "\TANF Signature Notification.doc"
    
    Call CreateMemoFromTemplate(ws, programType, templatePath, outputPath, "TANF Signature")
    
    MsgBox "TANF Signature Notification saved to: " & outputPath, vbInformation, "Success"
    
    Exit Sub
    
ErrorHandler:
    LogError "GenerateTANFSignatureNotification", Err.Number, Err.Description, ""
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub

' ----------------------------------------------------------------------------
' Sub: GenerateMAZeroIncome
' ----------------------------------------------------------------------------
' PURPOSE:
'   Generates MA Zero Income verification request memo.
'
' V1 EQUIVALENT: MA_Zero()
' ----------------------------------------------------------------------------
Public Sub GenerateMAZeroIncome()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim programType As ProgramType
    Dim dqcPath As String
    Dim templatePath As String
    Dim outputPath As String
    
    Set ws = ActiveSheet
    programType = GetProgramFromSheetName(ws.Name)
    
    If programType <> PROG_MA_POS And programType <> PROG_MA_NEG Then
        MsgBox "MA Zero Income is only available for MA reviews.", _
               vbInformation, "Not Applicable"
        Exit Sub
    End If
    
    dqcPath = GetDQCDriveLetterOrError()
    templatePath = dqcPath & "Finding Memo\MA Zero Income Request.docx"
    
    If Not PathExists(templatePath) Then
        MsgBox "Template not found at: " & templatePath, vbCritical, "Template Not Found"
        Exit Sub
    End If
    
    outputPath = ActiveWorkbook.Path & "\MA Zero Income Request.doc"
    
    Call CreateMemoFromTemplate(ws, programType, templatePath, outputPath, "MA Zero Income")
    
    MsgBox "MA Zero Income Request saved to: " & outputPath, vbInformation, "Success"
    
    Exit Sub
    
ErrorHandler:
    LogError "GenerateMAZeroIncome", Err.Number, Err.Description, ""
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub

' ----------------------------------------------------------------------------
' Sub: GeneratePendingLetter
' ----------------------------------------------------------------------------
' PURPOSE:
'   Generates PA472 Pending Letter (general pending verification request).
'
' V1 EQUIVALENT: PendLet()
' ----------------------------------------------------------------------------
Public Sub GeneratePendingLetter()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim programType As ProgramType
    Dim dqcPath As String
    Dim templatePath As String
    Dim outputPath As String
    
    Set ws = ActiveSheet
    programType = GetProgramFromSheetName(ws.Name)
    
    dqcPath = GetDQCDriveLetterOrError()
    templatePath = dqcPath & "Finding Memo\PA472 Pending Letter.docx"
    
    If Not PathExists(templatePath) Then
        MsgBox "Template not found at: " & templatePath, vbCritical, "Template Not Found"
        Exit Sub
    End If
    
    outputPath = ActiveWorkbook.Path & "\PA472 Pending Letter.doc"
    
    Call CreateMemoFromTemplate(ws, programType, templatePath, outputPath, "PA472 Pending")
    
    MsgBox "PA472 Pending Letter saved to: " & outputPath, vbInformation, "Success"
    
    Exit Sub
    
ErrorHandler:
    LogError "GeneratePendingLetter", Err.Number, Err.Description, ""
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub

' ----------------------------------------------------------------------------
' Sub: GenerateSelfEmploymentBasic
' ----------------------------------------------------------------------------
' PURPOSE:
'   Generates Basic Self-Employment verification form.
'
' V1 EQUIVALENT: SelfEmp()
' ----------------------------------------------------------------------------
Public Sub GenerateSelfEmploymentBasic()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim programType As ProgramType
    Dim dqcPath As String
    Dim templatePath As String
    Dim outputPath As String
    
    Set ws = ActiveSheet
    programType = GetProgramFromSheetName(ws.Name)
    
    dqcPath = GetDQCDriveLetterOrError()
    templatePath = dqcPath & "Finding Memo\Self Employment Basic.docx"
    
    If Not PathExists(templatePath) Then
        MsgBox "Template not found at: " & templatePath, vbCritical, "Template Not Found"
        Exit Sub
    End If
    
    outputPath = ActiveWorkbook.Path & "\Self Employment Basic Form.doc"
    
    Call CreateMemoFromTemplate(ws, programType, templatePath, outputPath, "Self Employment Basic")
    
    MsgBox "Self Employment Basic Form saved to: " & outputPath, vbInformation, "Success"
    
    Exit Sub
    
ErrorHandler:
    LogError "GenerateSelfEmploymentBasic", Err.Number, Err.Description, ""
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub

' ----------------------------------------------------------------------------
' Sub: GenerateSelfEmploymentDetailed
' ----------------------------------------------------------------------------
' PURPOSE:
'   Generates Detailed Self-Employment verification form.
'
' V1 EQUIVALENT: SelfEmpDet()
' ----------------------------------------------------------------------------
Public Sub GenerateSelfEmploymentDetailed()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim programType As ProgramType
    Dim dqcPath As String
    Dim templatePath As String
    Dim outputPath As String
    
    Set ws = ActiveSheet
    programType = GetProgramFromSheetName(ws.Name)
    
    dqcPath = GetDQCDriveLetterOrError()
    templatePath = dqcPath & "Finding Memo\Self Employment Detailed.docx"
    
    If Not PathExists(templatePath) Then
        MsgBox "Template not found at: " & templatePath, vbCritical, "Template Not Found"
        Exit Sub
    End If
    
    outputPath = ActiveWorkbook.Path & "\Self Employment Detailed Form.doc"
    
    Call CreateMemoFromTemplate(ws, programType, templatePath, outputPath, "Self Employment Detailed")
    
    MsgBox "Self Employment Detailed Form saved to: " & outputPath, vbInformation, "Success"
    
    Exit Sub
    
ErrorHandler:
    LogError "GenerateSelfEmploymentDetailed", Err.Number, Err.Description, ""
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub

' ----------------------------------------------------------------------------
' Sub: GenerateMASupportForm
' ----------------------------------------------------------------------------
' PURPOSE:
'   Generates MA Support Form (PA76).
'
' V1 EQUIVALENT: PA76()
' ----------------------------------------------------------------------------
Public Sub GenerateMASupportForm()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim programType As ProgramType
    Dim dqcPath As String
    Dim templatePath As String
    Dim outputPath As String
    
    Set ws = ActiveSheet
    programType = GetProgramFromSheetName(ws.Name)
    
    dqcPath = GetDQCDriveLetterOrError()
    templatePath = dqcPath & "Finding Memo\MA Support Form.docx"
    
    If Not PathExists(templatePath) Then
        MsgBox "Template not found at: " & templatePath, vbCritical, "Template Not Found"
        Exit Sub
    End If
    
    outputPath = ActiveWorkbook.Path & "\MA Support Form.doc"
    
    Call CreateMemoFromTemplate(ws, programType, templatePath, outputPath, "MA Support")
    
    MsgBox "MA Support Form saved to: " & outputPath, vbInformation, "Success"
    
    Exit Sub
    
ErrorHandler:
    LogError "GenerateMASupportForm", Err.Number, Err.Description, ""
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub

' ----------------------------------------------------------------------------
' Sub: GenerateHouseholdComposition
' ----------------------------------------------------------------------------
' PURPOSE:
'   Generates Household Composition verification form.
'
' V1 EQUIVALENT: PA78()
' ----------------------------------------------------------------------------
Public Sub GenerateHouseholdComposition()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim programType As ProgramType
    Dim dqcPath As String
    Dim templatePath As String
    Dim outputPath As String
    
    Set ws = ActiveSheet
    programType = GetProgramFromSheetName(ws.Name)
    
    dqcPath = GetDQCDriveLetterOrError()
    templatePath = dqcPath & "Finding Memo\Household Composition.docx"
    
    If Not PathExists(templatePath) Then
        MsgBox "Template not found at: " & templatePath, vbCritical, "Template Not Found"
        Exit Sub
    End If
    
    outputPath = ActiveWorkbook.Path & "\Household Composition Form.doc"
    
    Call CreateMemoFromTemplate(ws, programType, templatePath, outputPath, "Household Comp")
    
    MsgBox "Household Composition Form saved to: " & outputPath, vbInformation, "Success"
    
    Exit Sub
    
ErrorHandler:
    LogError "GenerateHouseholdComposition", Err.Number, Err.Description, ""
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub

' ----------------------------------------------------------------------------
' Sub: GenerateThresholdMemo
' ----------------------------------------------------------------------------
' PURPOSE:
'   Generates Error Under Threshold memo when error amount is below reporting.
'
' V1 EQUIVALENT: Threshold()
' ----------------------------------------------------------------------------
Public Sub GenerateThresholdMemo()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim programType As ProgramType
    Dim dqcPath As String
    Dim templatePath As String
    Dim outputPath As String
    
    Set ws = ActiveSheet
    programType = GetProgramFromSheetName(ws.Name)
    
    dqcPath = GetDQCDriveLetterOrError()
    templatePath = dqcPath & "Finding Memo\Error Under Threshold.docx"
    
    If Not PathExists(templatePath) Then
        MsgBox "Template not found at: " & templatePath, vbCritical, "Template Not Found"
        Exit Sub
    End If
    
    outputPath = ActiveWorkbook.Path & "\Error Under Threshold Memo.doc"
    
    Call CreateMemoFromTemplate(ws, programType, templatePath, outputPath, "Threshold")
    
    MsgBox "Error Under Threshold Memo saved to: " & outputPath, vbInformation, "Success"
    
    Exit Sub
    
ErrorHandler:
    LogError "GenerateThresholdMemo", Err.Number, Err.Description, ""
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub

' ----------------------------------------------------------------------------
' Sub: GenerateMASAVE
' ----------------------------------------------------------------------------
' PURPOSE:
'   Generates MA SAVE Deficiency memo for citizenship verification.
'
' V1 EQUIVALENT: MA_SAVE()
' ----------------------------------------------------------------------------
Public Sub GenerateMASAVE()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim programType As ProgramType
    Dim dqcPath As String
    Dim templatePath As String
    Dim outputPath As String
    
    Set ws = ActiveSheet
    programType = GetProgramFromSheetName(ws.Name)
    
    If programType <> PROG_MA_POS And programType <> PROG_MA_NEG Then
        MsgBox "MA SAVE is only available for MA reviews.", _
               vbInformation, "Not Applicable"
        Exit Sub
    End If
    
    dqcPath = GetDQCDriveLetterOrError()
    templatePath = dqcPath & "Finding Memo\MA SAVE Deficiency.docx"
    
    If Not PathExists(templatePath) Then
        MsgBox "Template not found at: " & templatePath, vbCritical, "Template Not Found"
        Exit Sub
    End If
    
    outputPath = ActiveWorkbook.Path & "\MA SAVE Deficiency.doc"
    
    Call CreateMemoFromTemplate(ws, programType, templatePath, outputPath, "MA SAVE")
    
    MsgBox "MA SAVE Deficiency memo saved to: " & outputPath, vbInformation, "Success"
    
    Exit Sub
    
ErrorHandler:
    LogError "GenerateMASAVE", Err.Number, Err.Description, ""
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub

' ----------------------------------------------------------------------------
' Sub: GenerateFuneralHomeLetter
' ----------------------------------------------------------------------------
' PURPOSE:
'   Generates Funeral Home verification letter.
'
' V1 EQUIVALENT: Funeral()
' ----------------------------------------------------------------------------
Public Sub GenerateFuneralHomeLetter()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim programType As ProgramType
    Dim dqcPath As String
    Dim templatePath As String
    Dim outputPath As String
    
    Set ws = ActiveSheet
    programType = GetProgramFromSheetName(ws.Name)
    
    dqcPath = GetDQCDriveLetterOrError()
    templatePath = dqcPath & "Finding Memo\Funeral Home Letter.docx"
    
    If Not PathExists(templatePath) Then
        MsgBox "Template not found at: " & templatePath, vbCritical, "Template Not Found"
        Exit Sub
    End If
    
    outputPath = ActiveWorkbook.Path & "\Funeral Home Letter.doc"
    
    Call CreateMemoFromTemplate(ws, programType, templatePath, outputPath, "Funeral Home")
    
    MsgBox "Funeral Home Letter saved to: " & outputPath, vbInformation, "Success"
    
    Exit Sub
    
ErrorHandler:
    LogError "GenerateFuneralHomeLetter", Err.Number, Err.Description, ""
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub

' ----------------------------------------------------------------------------
' Sub: GenerateAdoptionFosterCare
' ----------------------------------------------------------------------------
' PURPOSE:
'   Generates Adoption/Foster Care verification request.
'
' V1 EQUIVALENT: Adoption()
' ----------------------------------------------------------------------------
Public Sub GenerateAdoptionFosterCare()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim programType As ProgramType
    Dim dqcPath As String
    Dim templatePath As String
    Dim outputPath As String
    
    Set ws = ActiveSheet
    programType = GetProgramFromSheetName(ws.Name)
    
    dqcPath = GetDQCDriveLetterOrError()
    templatePath = dqcPath & "Finding Memo\Adoption Foster Care.docx"
    
    If Not PathExists(templatePath) Then
        MsgBox "Template not found at: " & templatePath, vbCritical, "Template Not Found"
        Exit Sub
    End If
    
    outputPath = ActiveWorkbook.Path & "\Adoption Foster Care Request.doc"
    
    Call CreateMemoFromTemplate(ws, programType, templatePath, outputPath, "Adoption Foster Care")
    
    MsgBox "Adoption/Foster Care Request saved to: " & outputPath, vbInformation, "Success"
    
    Exit Sub
    
ErrorHandler:
    LogError "GenerateAdoptionFosterCare", Err.Number, Err.Description, ""
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub

' ----------------------------------------------------------------------------
' Sub: GenerateTANFSAVE
' ----------------------------------------------------------------------------
' PURPOSE:
'   Generates TANF SAVE verification request.
'
' V1 EQUIVALENT: TANF_SAVE()
' ----------------------------------------------------------------------------
Public Sub GenerateTANFSAVE()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim programType As ProgramType
    Dim dqcPath As String
    Dim templatePath As String
    Dim outputPath As String
    
    Set ws = ActiveSheet
    programType = GetProgramFromSheetName(ws.Name)
    
    If programType <> PROG_TANF Then
        MsgBox "TANF SAVE is only available for TANF reviews.", _
               vbInformation, "Not Applicable"
        Exit Sub
    End If
    
    dqcPath = GetDQCDriveLetterOrError()
    templatePath = dqcPath & "Finding Memo\TANF SAVE.docx"
    
    If Not PathExists(templatePath) Then
        MsgBox "Template not found at: " & templatePath, vbCritical, "Template Not Found"
        Exit Sub
    End If
    
    outputPath = ActiveWorkbook.Path & "\TANF SAVE Request.doc"
    
    Call CreateMemoFromTemplate(ws, programType, templatePath, outputPath, "TANF SAVE")
    
    MsgBox "TANF SAVE Request saved to: " & outputPath, vbInformation, "Success"
    
    Exit Sub
    
ErrorHandler:
    LogError "GenerateTANFSAVE", Err.Number, Err.Description, ""
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub

' ----------------------------------------------------------------------------
' Sub: GenerateLEP
' ----------------------------------------------------------------------------
' PURPOSE:
'   Generates LEP (Limited English Proficiency) document.
'
' V1 EQUIVALENT: LEP()
' ----------------------------------------------------------------------------
Public Sub GenerateLEP()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim programType As ProgramType
    Dim dqcPath As String
    Dim templatePath As String
    Dim outputPath As String
    
    Set ws = ActiveSheet
    programType = GetProgramFromSheetName(ws.Name)
    
    dqcPath = GetDQCDriveLetterOrError()
    templatePath = dqcPath & "Finding Memo\LEP Document.docx"
    
    If Not PathExists(templatePath) Then
        MsgBox "Template not found at: " & templatePath, vbCritical, "Template Not Found"
        Exit Sub
    End If
    
    outputPath = ActiveWorkbook.Path & "\LEP Document.doc"
    
    Call CreateMemoFromTemplate(ws, programType, templatePath, outputPath, "LEP")
    
    MsgBox "LEP Document saved to: " & outputPath, vbInformation, "Success"
    
    Exit Sub
    
ErrorHandler:
    LogError "GenerateLEP", Err.Number, Err.Description, ""
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub

' ----------------------------------------------------------------------------
' Sub: GenerateTANFPending
' ----------------------------------------------------------------------------
' PURPOSE:
'   Generates TANF Pending verification request.
'
' V1 EQUIVALENT: TANFPend()
' ----------------------------------------------------------------------------
Public Sub GenerateTANFPending()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim programType As ProgramType
    Dim dqcPath As String
    Dim templatePath As String
    Dim outputPath As String
    
    Set ws = ActiveSheet
    programType = GetProgramFromSheetName(ws.Name)
    
    If programType <> PROG_TANF Then
        MsgBox "TANF Pending is only available for TANF reviews.", _
               vbInformation, "Not Applicable"
        Exit Sub
    End If
    
    dqcPath = GetDQCDriveLetterOrError()
    templatePath = dqcPath & "Finding Memo\TANF Pending.docx"
    
    If Not PathExists(templatePath) Then
        MsgBox "Template not found at: " & templatePath, vbCritical, "Template Not Found"
        Exit Sub
    End If
    
    outputPath = ActiveWorkbook.Path & "\TANF Pending Letter.doc"
    
    Call CreateMemoFromTemplate(ws, programType, templatePath, outputPath, "TANF Pending")
    
    MsgBox "TANF Pending Letter saved to: " & outputPath, vbInformation, "Success"
    
    Exit Sub
    
ErrorHandler:
    LogError "GenerateTANFPending", Err.Number, Err.Description, ""
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub

' ----------------------------------------------------------------------------
' Sub: GenerateSchoolVerification
' ----------------------------------------------------------------------------
' PURPOSE:
'   Generates School Verification request letter.
'
' V1 EQUIVALENT: TANFSchool()
' ----------------------------------------------------------------------------
Public Sub GenerateSchoolVerification()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim programType As ProgramType
    Dim dqcPath As String
    Dim templatePath As String
    Dim outputPath As String
    
    Set ws = ActiveSheet
    programType = GetProgramFromSheetName(ws.Name)
    
    dqcPath = GetDQCDriveLetterOrError()
    templatePath = dqcPath & "Finding Memo\School Verification.docx"
    
    If Not PathExists(templatePath) Then
        MsgBox "Template not found at: " & templatePath, vbCritical, "Template Not Found"
        Exit Sub
    End If
    
    outputPath = ActiveWorkbook.Path & "\School Verification Request.doc"
    
    Call CreateMemoFromTemplate(ws, programType, templatePath, outputPath, "School Verification")
    
    MsgBox "School Verification Request saved to: " & outputPath, vbInformation, "Success"
    
    Exit Sub
    
ErrorHandler:
    LogError "GenerateSchoolVerification", Err.Number, Err.Description, ""
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub

' ----------------------------------------------------------------------------
' Sub: GenerateMAWaiverMemo
' ----------------------------------------------------------------------------
' PURPOSE:
'   Generates MA Waiver Memo.
' ----------------------------------------------------------------------------
Public Sub GenerateMAWaiverMemo()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim programType As ProgramType
    Dim dqcPath As String
    Dim templatePath As String
    Dim outputPath As String
    
    Set ws = ActiveSheet
    programType = GetProgramFromSheetName(ws.Name)
    
    dqcPath = GetDQCDriveLetterOrError()
    templatePath = dqcPath & "Finding Memo\MA Waiver Memo.docx"
    
    If Not PathExists(templatePath) Then
        MsgBox "Template not found at: " & templatePath, vbCritical, "Template Not Found"
        Exit Sub
    End If
    
    outputPath = ActiveWorkbook.Path & "\MA Waiver Memo.doc"
    
    Call CreateMemoFromTemplate(ws, programType, templatePath, outputPath, "MA Waiver")
    
    MsgBox "MA Waiver Memo saved to: " & outputPath, vbInformation, "Success"
    
    Exit Sub
    
ErrorHandler:
    LogError "GenerateMAWaiverMemo", Err.Number, Err.Description, ""
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub

' ----------------------------------------------------------------------------
' Sub: GenerateMALTCWaiverMemo
' ----------------------------------------------------------------------------
' PURPOSE:
'   Generates MA LTC (Long Term Care) Waiver Memo.
' ----------------------------------------------------------------------------
Public Sub GenerateMALTCWaiverMemo()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim programType As ProgramType
    Dim dqcPath As String
    Dim templatePath As String
    Dim outputPath As String
    
    Set ws = ActiveSheet
    programType = GetProgramFromSheetName(ws.Name)
    
    dqcPath = GetDQCDriveLetterOrError()
    templatePath = dqcPath & "Finding Memo\MA LTC Waiver Memo.docx"
    
    If Not PathExists(templatePath) Then
        MsgBox "Template not found at: " & templatePath, vbCritical, "Template Not Found"
        Exit Sub
    End If
    
    outputPath = ActiveWorkbook.Path & "\MA LTC Waiver Memo.doc"
    
    Call CreateMemoFromTemplate(ws, programType, templatePath, outputPath, "MA LTC Waiver")
    
    MsgBox "MA LTC Waiver Memo saved to: " & outputPath, vbInformation, "Success"
    
    Exit Sub
    
ErrorHandler:
    LogError "GenerateMALTCWaiverMemo", Err.Number, Err.Description, ""
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub


' ============================================================================
' HELPER FUNCTIONS
' ============================================================================

' ----------------------------------------------------------------------------
' Function: GetReviewNumber
' ----------------------------------------------------------------------------
' PURPOSE:
'   Extracts the review number from the schedule based on program type.
' ----------------------------------------------------------------------------
Private Function GetReviewNumber(ByRef ws As Worksheet, _
                                 ByVal programType As ProgramType) As String
    On Error Resume Next
    
    Select Case programType
        Case PROG_TANF, PROG_GA
            GetReviewNumber = CStr(ws.Range("A10").Value)
        Case PROG_SNAP_POS
            GetReviewNumber = CStr(ws.Range("A18").Value)
        Case PROG_SNAP_NEG
            GetReviewNumber = CStr(ws.Range("A18").Value)
        Case PROG_MA_POS, PROG_MA_NEG, PROG_MA_PE
            GetReviewNumber = CStr(ws.Range("D2").Value)
        Case Else
            GetReviewNumber = ws.Name
    End Select
    
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' Function: GetSampleMonth
' ----------------------------------------------------------------------------
' PURPOSE:
'   Extracts the sample month from the schedule based on program type.
' ----------------------------------------------------------------------------
Private Function GetSampleMonth(ByRef ws As Worksheet, _
                                ByVal programType As ProgramType) As String
    On Error Resume Next
    
    Select Case programType
        Case PROG_TANF, PROG_GA
            GetSampleMonth = CStr(ws.Range("AB10").Value)
        Case PROG_SNAP_POS
            GetSampleMonth = CStr(ws.Range("AD18").Value) & CStr(ws.Range("AG18").Value)
        Case PROG_SNAP_NEG
            GetSampleMonth = CStr(ws.Range("C18").Value) & CStr(ws.Range("G18").Value)
        Case PROG_MA_POS, PROG_MA_NEG, PROG_MA_PE
            GetSampleMonth = CStr(ws.Range("E2").Value)
        Case Else
            GetSampleMonth = Format(Date, "MMYYYY")
    End Select
    
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' Sub: CreateMemoFromTemplate
' ----------------------------------------------------------------------------
' PURPOSE:
'   Core function that performs mail merge using Word template.
'   Extracts data from the schedule and populates the template.
' ----------------------------------------------------------------------------
Private Sub CreateMemoFromTemplate(ByRef ws As Worksheet, _
                                   ByVal programType As ProgramType, _
                                   ByVal templatePath As String, _
                                   ByVal outputPath As String, _
                                   ByVal memoType As String)
    On Error GoTo ErrorHandler
    
    Dim wdApp As Object
    Dim wdDoc As Object
    Dim dqcPath As String
    Dim dsPath As String
    Dim dsWB As Workbook
    Dim dsWS As Worksheet
    
    Application.ScreenUpdating = False
    Application.StatusBar = "Creating " & memoType & " memo..."
    
    ' Get DQC path
    dqcPath = GetDQCDriveLetterOrError()
    
    ' Open and populate data source
    dsPath = dqcPath & "Finding Memo\Finding Memo Data Source.xlsx"
    
    If PathExists(dsPath) Then
        ' Copy data source to temp file
        Dim tempDSPath As String
        tempDSPath = ActiveWorkbook.Path & "\FM DS Temp.xlsx"
        FileCopy dsPath, tempDSPath
        
        ' Open temp data source
        Set dsWB = Workbooks.Open(Filename:=tempDSPath, UpdateLinks:=False)
        Set dsWS = dsWB.Worksheets(1)
        
        ' Populate data source based on program type
        Call PopulateMemoDataSource(ws, dsWS, programType)
        
        ' Save and close data source
        dsWB.Save
        dsWB.Close
    End If
    
    ' Copy template to temp file
    Dim tempTemplatePath As String
    tempTemplatePath = ActiveWorkbook.Path & "\FM temp.docx"
    FileCopy templatePath, tempTemplatePath
    
    ' Create Word application
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = True
    
    ' Open template and execute mail merge
    Set wdDoc = wdApp.Documents.Open(Filename:=tempTemplatePath)
    
    If PathExists(tempDSPath) Then
        With wdDoc.MailMerge
            .OpenDataSource Name:=tempDSPath, _
                ReadOnly:=True, LinkToSource:=0, AddToRecentFiles:=False
            .Destination = 0  ' wdSendToNewDocument
            .SuppressBlankLines = True
            .Execute Pause:=False
        End With
    End If
    
    Application.StatusBar = "Saving document..."
    
    ' Save merged document
    wdApp.ActiveDocument.Content.Font.Name = "Arial"
    wdApp.ActiveDocument.SaveAs outputPath
    wdApp.ActiveDocument.Close
    
    ' Close template
    wdDoc.Close SaveChanges:=False
    
    ' Cleanup
    wdApp.Quit
    Set wdApp = Nothing
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    On Error Resume Next
    If Not wdDoc Is Nothing Then wdDoc.Close SaveChanges:=False
    If Not wdApp Is Nothing Then wdApp.Quit
    Set wdApp = Nothing
    
    LogError "CreateMemoFromTemplate", Err.Number, Err.Description, memoType
    Err.Raise Err.Number, , Err.Description
End Sub

' ----------------------------------------------------------------------------
' Sub: PopulateMemoDataSource
' ----------------------------------------------------------------------------
' PURPOSE:
'   Populates the memo data source spreadsheet with case information.
' ----------------------------------------------------------------------------
Private Sub PopulateMemoDataSource(ByRef schedWS As Worksheet, _
                                    ByRef dsWS As Worksheet, _
                                    ByVal programType As ProgramType)
    On Error Resume Next
    
    ' Clear existing data
    dsWS.Range("B2:E2").Value = ""
    dsWS.Range("G2").Value = ""
    dsWS.Range("O2:S2").Value = ""
    dsWS.Range("W2:Y2").Value = ""
    
    Select Case programType
        Case PROG_TANF, PROG_GA
            Call PopulateTANFGADataSource(schedWS, dsWS, programType)
        Case PROG_SNAP_POS
            Call PopulateSNAPPosDataSource(schedWS, dsWS)
        Case PROG_SNAP_NEG
            Call PopulateSNAPNegDataSource(schedWS, dsWS)
        Case PROG_MA_POS, PROG_MA_NEG, PROG_MA_PE
            Call PopulateMADataSource(schedWS, dsWS)
    End Select
    
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' Sub: PopulateTANFGADataSource
' ----------------------------------------------------------------------------
Private Sub PopulateTANFGADataSource(ByRef schedWS As Worksheet, _
                                      ByRef dsWS As Worksheet, _
                                      ByVal programType As ProgramType)
    On Error Resume Next
    
    Dim thiswb As Workbook
    Dim compSheetName As String
    
    Set thiswb = schedWS.Parent
    
    ' Client Information
    dsWS.Range("B2").Value = StrConv(schedWS.Range("B2").Value, vbProperCase)  ' Client Name
    dsWS.Range("C2").Value = Right(schedWS.Range("U10").Value, 2) & "/" & _
                              schedWS.Range("I10").Value  ' County/Case Number
    dsWS.Range("D2").Value = schedWS.Range("A10").Value  ' Review Number
    dsWS.Range("E2").Value = Left(schedWS.Range("AB10").Value, 2) & "/" & _
                              Right(schedWS.Range("AB10").Value, 4)  ' Review Month
    
    ' Reviewer info
    dsWS.Range("A7").Value = Val(schedWS.Range("AO3").Value & schedWS.Range("AP3").Value)
    dsWS.Range("G2").Value = dsWS.Range("C7").Value & "/" & schedWS.Range("AG2").Value
    
    ' Error info
    dsWS.Range("R2").Value = schedWS.Range("J61").Value & " - " & _
                              schedWS.Range("O61").Value & " - " & _
                              schedWS.Range("T61").Value
    dsWS.Range("E17").Value = Val(schedWS.Range("AL10").Value)
    dsWS.Range("Q2").Value = schedWS.Range("AO10").Value  ' Error Amount
    
    ' Benefit amounts from computation sheet
    If programType = PROG_TANF Then
        compSheetName = "TANF Computation"
        dsWS.Range("O2").Value = thiswb.Sheets(compSheetName).Range("B72").Value
    ElseIf programType = PROG_GA Then
        compSheetName = "GA Computation"
        dsWS.Range("O2").Value = thiswb.Sheets(compSheetName).Range("B73").Value
    End If
    
    ' Address
    dsWS.Range("Z2").Value = StrConv(schedWS.Range("B3").Value, vbProperCase)
    
    Dim temparry() As String
    Dim tempStr As String
    temparry = Split(schedWS.Range("B4").Value, ",")
    tempStr = StrConv(Trim(temparry(LBound(temparry))), vbProperCase)
    dsWS.Range("AA2").Value = tempStr & ", " & Trim(temparry(UBound(temparry)))
    
    ' CAO
    If programType = PROG_TANF Then
        dsWS.Range("AK2").Value = schedWS.Range("O4").Value & " Assistance Office"
    Else
        dsWS.Range("AK2").Value = schedWS.Range("M5").Value & " Assistance Office"
    End If
    
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' Sub: PopulateSNAPPosDataSource
' ----------------------------------------------------------------------------
Private Sub PopulateSNAPPosDataSource(ByRef schedWS As Worksheet, _
                                       ByRef dsWS As Worksheet)
    On Error Resume Next
    
    ' Client Information
    dsWS.Range("B2").Value = StrConv(schedWS.Range("B4").Value, vbProperCase)
    dsWS.Range("C2").Value = schedWS.Range("C155").Value & schedWS.Range("D155").Value & _
                              "/" & Left(schedWS.Range("I18").Value, 9)
    dsWS.Range("D2").Value = schedWS.Range("A18").Value
    dsWS.Range("E2").Value = schedWS.Range("AD18").Value & "/" & schedWS.Range("AG18").Value
    
    ' CAO
    dsWS.Range("AK2").Value = schedWS.Range("M5").Value & " Assistance Office"
    
    ' Reviewer info
    dsWS.Range("A7").Value = Val(schedWS.Range("U5").Value & schedWS.Range("V5").Value)
    dsWS.Range("G2").Value = dsWS.Range("C7").Value & "/" & schedWS.Range("AG2").Value
    
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' Sub: PopulateSNAPNegDataSource
' ----------------------------------------------------------------------------
Private Sub PopulateSNAPNegDataSource(ByRef schedWS As Worksheet, _
                                       ByRef dsWS As Worksheet)
    On Error Resume Next
    
    ' Client Information
    dsWS.Range("B2").Value = StrConv(schedWS.Range("B4").Value, vbProperCase)
    dsWS.Range("C2").Value = schedWS.Range("AA5").Value & "/" & _
                              Left(schedWS.Range("I18").Value, 9)
    dsWS.Range("D2").Value = schedWS.Range("A18").Value
    dsWS.Range("E2").Value = schedWS.Range("C18").Value & "/" & schedWS.Range("G18").Value
    
    ' CAO
    dsWS.Range("AK2").Value = schedWS.Range("M5").Value & " Assistance Office"
    
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' Sub: PopulateMADataSource
' ----------------------------------------------------------------------------
Private Sub PopulateMADataSource(ByRef schedWS As Worksheet, _
                                  ByRef dsWS As Worksheet)
    On Error Resume Next
    
    ' MA uses different cell locations
    dsWS.Range("B2").Value = StrConv(schedWS.Range("A13").Value, vbProperCase)
    dsWS.Range("D2").Value = schedWS.Range("D2").Value
    dsWS.Range("E2").Value = schedWS.Range("E2").Value
    
    On Error GoTo 0
End Sub


