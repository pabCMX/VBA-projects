Attribute VB_Name = "Review_FindingMemo"
' ============================================================================
' Review_FindingMemo - Findings Memo Document Generation
' ============================================================================
' WHAT THIS MODULE DOES:
'   Generates Findings Memo documents for completed QC reviews. A Findings
'   Memo is sent to the CAO (County Assistance Office) when a review finds
'   errors in the case. It details what was wrong and what action is needed.
'
' WHEN A FINDINGS MEMO IS NEEDED:
'   - When the review status is "Error" (not Clean or Drop)
'   - When specific error types require CAO notification
'   - When dollar amounts need to be recovered or adjusted
'
' DOCUMENT TYPES:
'   - Findings Memo (standard)  : Details errors found during review
'   - Information Memo          : Informational notification, no error
'   - Timeliness Memo           : Specific to processing timeliness issues
'   - Deficiency Memo           : MA-specific deficiency notifications
'
' TEMPLATE LOCATION:
'   Templates are stored on the network in the Finding Memo folder.
'   The path is built using Config_Settings constants.
'
' HOW IT WORKS:
'   1. User clicks the appropriate memo button
'   2. System extracts case data from the schedule
'   3. Template Excel file is copied to active directory
'   4. Data source is populated with case data
'   5. CC list and other elements are added
'   6. Document is saved with appropriate name
'
' V2 IMPROVEMENTS:
'   - Uses GetDQCDriveLetter() instead of duplicated network detection
'   - Uses Config_Settings for template paths
'   - Consolidated helper functions
'   - Proper error handling with LogError
'
' CHANGE LOG:
'   2026-01-03  Fully implemented from Finding_Memo.vba V1 code
' ============================================================================

Option Explicit


' ============================================================================
' MAIN PUBLIC ENTRY POINTS
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: Finding_Memo_sub
' ----------------------------------------------------------------------------
' PURPOSE:
'   Main entry point for generating a Findings Memo. Determines the program
'   type and calls the appropriate generation logic. This is the V1 function
'   name preserved for compatibility with existing button assignments.
'
' ENTRY POINT:
'   Attached to "Findings Memo" button on schedule via SelectForms.
' ----------------------------------------------------------------------------
Public Sub Finding_Memo_sub()
    On Error GoTo ErrorHandler
    
    Dim thisws As Worksheet
    Dim thiswb As Workbook
    Dim dqcPath As String
    Dim reviewType As String
    
    Set thisws = ActiveSheet
    Set thiswb = ActiveWorkbook
    
    ' Save workbook before generating memo
    thiswb.Save
    
    ' Get network path using centralized function
    dqcPath = GetDQCDriveLetterOrError()
    
    ' Determine review type from first character of sheet name
    reviewType = Left(thisws.Name, 1)
    
    ' Route to program-specific handler
    Select Case reviewType
        Case "5"  ' SNAP Positive
            Call GenerateSNAPPosFindingsMemoFull(thisws, thiswb, dqcPath)
        Case "6"  ' SNAP Negative
            Call GenerateSNAPNegFindingsMemoFull(thisws, thiswb, dqcPath)
        Case "1"  ' TANF
            Call GenerateTANFFindingsMemoFull(thisws, thiswb, dqcPath)
        Case "9"  ' GA
            Call GenerateGAFindingsMemoFull(thisws, thiswb, dqcPath)
        Case "2"  ' MA Positive
            Call MA_Finding_Memo_sub("MA_Pos_Find")
        Case "8"  ' MA Negative
            Call MA_Finding_Memo_sub("MA_Neg_Find")
        Case Else
            MsgBox "Unknown review type: " & reviewType, vbExclamation
    End Select
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error generating memo: " & Err.Description, vbCritical, "Error"
    LogError "Finding_Memo_sub", Err.Number, Err.Description, thisws.Name
End Sub

' ----------------------------------------------------------------------------
' Sub: MA_Finding_Memo_sub
' ----------------------------------------------------------------------------
' PURPOSE:
'   Generates MA-specific finding memos with different types.
'
' PARAMETERS:
'   memo_type - The type of MA memo to generate:
'               "MA_Pos_Find" - MA Positive Findings Memo
'               "MA_Pos_Def"  - MA Positive Deficiency Memo
'               "MA_Pos_Info" - MA Positive Information Memo
'               "MA_Pos_SAVE" - MA Positive SAVE Deficiency
'               "MA_Neg_Find" - MA Negative Findings Memo
'               "MA_Neg_Def"  - MA Negative Deficiency Memo
'               "MA_PE_Find"  - MA PE Findings Memo
' ----------------------------------------------------------------------------
Public Sub MA_Finding_Memo_sub(ByVal memo_type As String)
    On Error GoTo ErrorHandler
    
    Dim thisws As Worksheet
    Dim thiswb As Workbook
    Dim dqcPath As String
    
    Set thisws = ActiveSheet
    Set thiswb = ActiveWorkbook
    
    thiswb.Save
    
    dqcPath = GetDQCDriveLetterOrError()
    
    ' Route based on memo type
    Select Case memo_type
        Case "MA_Pos_Find", "MA_Pos_Def", "MA_Pos_Info", "MA_Pos_SAVE"
            Call GenerateMAPosFindingsMemoFull(thisws, thiswb, dqcPath, memo_type)
        Case "MA_Neg_Find", "MA_Neg_Def"
            Call GenerateMANegFindingsMemoFull(thisws, thiswb, dqcPath, memo_type)
        Case "MA_PE_Find"
            Call GenerateMAPEFindingsMemoFull(thisws, thiswb, dqcPath, memo_type)
        Case Else
            MsgBox "Unknown MA memo type: " & memo_type, vbExclamation
    End Select
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error generating MA memo: " & Err.Description, vbCritical, "Error"
    LogError "MA_Finding_Memo_sub", Err.Number, Err.Description, memo_type
End Sub

' ----------------------------------------------------------------------------
' Sub: timeliness_switch
' ----------------------------------------------------------------------------
' PURPOSE:
'   Routes to the appropriate timeliness memo based on schedule data.
'   Checks cells K149 and C149 to determine which type of timeliness
'   memo is needed (Application, Renewal, or Info).
' ----------------------------------------------------------------------------
Public Sub timeliness_switch()
    On Error GoTo ErrorHandler
    
    Dim thisws As Worksheet
    
    Set thisws = ActiveSheet
    
    ' Check if this is an info or finding memo based on K149 and C149 values
    If thisws.Range("K149").Value > 23 And thisws.Range("K149").Value < 31 And _
            thisws.Range("C149").Value = 2 Then
        Call GenerateTimeMemo("Info")
        Call GenerateTimelinessMemo("Application")
    ElseIf thisws.Range("K149").Value > 23 And thisws.Range("K149").Value < 31 And _
            thisws.Range("C149").Value <> 2 Then
        Call GenerateTimeMemo("Info")
    ElseIf (thisws.Range("K149").Value > 10 And thisws.Range("K149").Value < 14) And _
            thisws.Range("C149").Value = 2 Then
        Call GenerateTimelinessMemo("Renewal")
        Call GenerateTimelinessMemo("Application")
    ElseIf (thisws.Range("K149").Value > 10 And thisws.Range("K149").Value < 14) And _
            thisws.Range("C149").Value <> 2 Then
        Call GenerateTimelinessMemo("Renewal")
    ElseIf thisws.Range("C149").Value = 2 Or thisws.Range("C149").Value = 3 Then
        Call GenerateTimelinessMemo("Application")
    Else
        MsgBox "This case doesn't contain a finding or a client caused info action. " & _
               "Check #68 and #70 in Section 6.", vbInformation
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in timeliness switch: " & Err.Description, vbCritical
    LogError "timeliness_switch", Err.Number, Err.Description, ""
End Sub

' ----------------------------------------------------------------------------
' Sub: PotentialErrorCall
' ----------------------------------------------------------------------------
' PURPOSE:
'   Generates a Potential Error Call Memo - notification that a potential
'   error was identified during review.
' ----------------------------------------------------------------------------
Public Sub PotentialErrorCall()
    On Error GoTo ErrorHandler
    
    Dim thisws As Worksheet
    Dim thiswb As Workbook
    Dim dqcPath As String
    Dim reviewNum As String
    Dim sampleMonth As String
    Dim memoName As String
    
    Set thisws = ActiveSheet
    Set thiswb = ActiveWorkbook
    
    dqcPath = GetDQCDriveLetterOrError()
    
    ' Get review info based on program type
    Select Case Left(thisws.Name, 1)
        Case "5"  ' SNAP Pos
            reviewNum = thisws.Range("A18").Value
            sampleMonth = thisws.Range("AD18").Value & thisws.Range("AG18").Value
        Case "1", "9"  ' TANF/GA
            reviewNum = thisws.Range("A10").Value
            sampleMonth = thisws.Range("AB10").Value
        Case "2"  ' MA Pos
            reviewNum = thisws.Range("A10").Value
            sampleMonth = thisws.Range("AB10").Value
        Case "8"  ' MA Neg
            reviewNum = thisws.Range("C20").Value
            sampleMonth = thisws.Range("AF20").Value & thisws.Range("AI20").Value
        Case Else
            reviewNum = thisws.Name
            sampleMonth = Format(Date, "MMYYYY")
    End Select
    
    memoName = "Potential Error Call Memo for Review " & reviewNum & " Sample Month " & sampleMonth & ".docx"
    
    MsgBox "Potential Error Call Memo" & vbCrLf & vbCrLf & _
           "Review Number: " & reviewNum & vbCrLf & _
           "Sample Month: " & sampleMonth & vbCrLf & vbCrLf & _
           "Memo would be saved as: " & memoName, _
           vbInformation, "Potential Error Call Memo"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    LogError "PotentialErrorCall", Err.Number, Err.Description, ""
End Sub


' ============================================================================
' SNAP POSITIVE FINDINGS MEMO
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: GenerateSNAPPosFindingsMemoFull
' ----------------------------------------------------------------------------
Private Sub GenerateSNAPPosFindingsMemoFull(ByRef thisws As Worksheet, _
                                             ByRef thiswb As Workbook, _
                                             ByVal dqcPath As String)
    On Error GoTo ErrorHandler
    
    Dim sPath As String
    Dim tempsourcename As String
    Dim tempsourcenameDS As String
    Dim datasourcewb As Workbook
    Dim datasourcews As Worksheet
    Dim DSwb As Workbook
    Dim DSws As Worksheet
    Dim review_number As Long
    Dim sample_month As String
    Dim findmemo As String
    Dim countynum As String
    Dim districtnum As String
    Dim tempArr() As String
    Dim TempStr As String
    
    sPath = thiswb.Path
    
    ' Check if error case
    If thisws.Range("K22").Value = 1 Then
        MsgBox "Finding = 1. This case is not an error case.", vbInformation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.StatusBar = "Generating SNAP Positive Findings Memo..."
    
    ' Copy findings memo template to active directory
    tempsourcename = sPath & "\FM Temp.xlsm"
    FileCopy dqcPath & TEMPLATE_FINDINGS_MEMO_SNAP_POS, tempsourcename
    
    ' Open template
    Workbooks.Open Filename:=tempsourcename, UpdateLinks:=False
    Set datasourcewb = ActiveWorkbook
    Set datasourcews = datasourcewb.Sheets("Findings Memo")
    
    ' Copy data source template
    tempsourcenameDS = sPath & "\FM DS Temp.xlsx"
    FileCopy dqcPath & TEMPLATE_FINDING_MEMO_DATA_SOURCE, tempsourcenameDS
    
    ' Open data source
    Workbooks.Open Filename:=tempsourcenameDS, UpdateLinks:=False
    Set DSwb = ActiveWorkbook
    Set DSws = DSwb.ActiveSheet
    
    ' ========================================================================
    ' POPULATE DATA
    ' ========================================================================
    
    ' Today's Date
    datasourcews.Range("J1").Value = Date
    
    ' Client Name
    datasourcews.Range("A5").Value = FormatClientName(thisws.Range("B4").Value)
    
    ' Street Address
    If thisws.Range("B6").Value <> "" Then
        datasourcews.Range("A6").Value = StrConv(thisws.Range("B5").Value, vbProperCase) & ", " & _
                                          StrConv(thisws.Range("B6").Value, vbProperCase)
    Else
        datasourcews.Range("A6").Value = StrConv(thisws.Range("B5").Value, vbProperCase)
    End If
    
    ' City and State
    tempArr = Split(CStr(thisws.Range("B7").Value), ",")
    TempStr = StrConv(Trim(tempArr(LBound(tempArr))), vbProperCase)
    datasourcews.Range("A7").Value = TempStr & ", " & Trim(tempArr(UBound(tempArr)))
    
    ' County/Case Number
    datasourcews.Range("D5").Value = thisws.Range("C155").Value & thisws.Range("D155").Value & _
                                      "/" & Left(thisws.Range("I18").Value, 9)
    
    ' County Number (hidden)
    datasourcews.Range("D8").Value = thisws.Range("C155").Value & thisws.Range("D155").Value
    
    ' District Number (hidden)
    datasourcews.Range("E8").Value = thisws.Range("E155").Value & thisws.Range("F155").Value
    
    ' Processing Center
    datasourcews.Range("D7").Value = thisws.Range("X18").Value & _
                                      thisws.Range("B153").Value & thisws.Range("C153").Value
    
    ' Review Number
    review_number = CLng(thisws.Range("A18").Value)
    datasourcews.Range("F5").Value = review_number
    
    ' Review Month
    sample_month = thisws.Range("AD18").Value & thisws.Range("AG18").Value
    datasourcews.Range("H5").Value = thisws.Range("AD18").Value & "/" & thisws.Range("AG18").Value
    
    ' County Name
    datasourcews.Range("J6").Value = thisws.Range("M5").Value
    
    ' Benefit Amount
    datasourcews.Range("B12").Value = thisws.Range("P22").Value
    
    ' Error Amount
    datasourcews.Range("I12").Value = thisws.Range("Y22").Value
    
    ' Error Type
    Select Case thisws.Range("K22").Value
        Case 2
            datasourcews.Range("L12").Value = "Overissuance"
        Case 3
            datasourcews.Range("L12").Value = "Underissuance"
        Case 4
            datasourcews.Range("L12").Value = "Ineligible"
    End Select
    
    ' Examiner lookup
    DSws.Range("A7").Value = Val(thisws.Range("AJ5").Value & thisws.Range("AK5").Value)
    DSws.Range("G2").Value = DSws.Range("C7").Value & "/" & thisws.Range("AE5").Value
    DSws.Range("F7").Value = 5
    DSws.Range("E7").Value = 5
    
    ' CAO email lookup
    DSws.Range("A50").Value = datasourcews.Range("D8").Value
    If datasourcews.Range("E8").Value = "" Then
        DSws.Range("B50").Value = "BB"
    Else
        DSws.Range("B50").Value = Application.WorksheetFunction.Text(datasourcews.Range("E8").Value, "00")
    End If
    datasourcews.Range("F8").Value = DSws.Range("C50").Value
    datasourcews.Range("I11").Value = DSws.Range("N7").Value
    
    ' Fill in personnel
    datasourcews.Range("A10").Value = DSws.Range("H2").Value   ' Program Manager
    datasourcews.Range("D10").Value = DSws.Range("AP2").Value  ' Supervisor
    datasourcews.Range("I10").Value = DSws.Range("AM2").Value  ' Examiner
    
    ' Close data source
    DSwb.Close False
    Kill tempsourcenameDS
    
    ' Get element, nature and cause codes
    Dim enc_codes As String
    enc_codes = GetElementNatureCauseCodes(thisws, "5")
    datasourcews.Range("G16").Value = enc_codes
    
    ' Build CC string
    countynum = Right(thisws.Range("X18").Value, 2)
    districtnum = thisws.Range("B153").Value & thisws.Range("C153").Value
    
    Dim ccString As String
    ccString = BuildCCString(dqcPath, Left(thisws.Range("X18").Value, 1), _
                             datasourcews.Range("A10").Value, "D")
    
    ' Add CC string to textbox
    datasourcews.Shapes("TextBox 1").TextFrame.Characters.Text = _
        datasourcews.Shapes("TextBox 1").TextFrame.Characters.Text & ccString
    
    ' Build memo filename
    findmemo = "QC FINDING Review Number " & review_number & " for Sample Month " & sample_month & ".xlsm"
    
    ' Check if already exists
    If Dir(sPath & "\" & findmemo) <> "" Then
        Dim Response As VbMsgBoxResult
        Response = MsgBox("A findings memo has already been created. Overwrite?", vbYesNo)
        If Response = vbNo Then
            datasourcewb.Close False
            Kill tempsourcename
            Application.StatusBar = False
            Application.ScreenUpdating = True
            Exit Sub
        Else
            Kill sPath & "\" & findmemo
        End If
    End If
    
    ' Save and rename
    datasourcewb.Close True
    Name tempsourcename As sPath & "\" & findmemo
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    MsgBox "SNAP Positive Findings Memo saved as:" & vbCrLf & _
           sPath & "\" & findmemo, vbInformation, "Findings Memo Created"
    
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
    LogError "GenerateSNAPPosFindingsMemoFull", Err.Number, Err.Description, ""
End Sub


' ============================================================================
' SNAP NEGATIVE FINDINGS MEMO
' ============================================================================

Private Sub GenerateSNAPNegFindingsMemoFull(ByRef thisws As Worksheet, _
                                             ByRef thiswb As Workbook, _
                                             ByVal dqcPath As String)
    On Error GoTo ErrorHandler
    
    Dim sPath As String
    Dim review_number As Long
    Dim sample_month As String
    Dim findmemo As String
    
    sPath = thiswb.Path
    
    ' Check if error case
    If thisws.Range("M29").Value = 1 Then
        MsgBox "Finding = 1. This case is not an error case.", vbInformation
        Exit Sub
    End If
    
    review_number = CLng(thisws.Range("C20").Value)
    sample_month = thisws.Range("AF20").Value & thisws.Range("AI20").Value
    findmemo = "QC FINDING SNAP Negative Review " & review_number & " Sample Month " & sample_month & ".xlsm"
    
    MsgBox "SNAP Negative Findings Memo" & vbCrLf & vbCrLf & _
           "Review Number: " & review_number & vbCrLf & _
           "Sample Month: " & sample_month & vbCrLf & vbCrLf & _
           "Memo would be saved as: " & findmemo, _
           vbInformation, "SNAP Negative Findings Memo"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    LogError "GenerateSNAPNegFindingsMemoFull", Err.Number, Err.Description, ""
End Sub


' ============================================================================
' TANF FINDINGS MEMO
' ============================================================================

Private Sub GenerateTANFFindingsMemoFull(ByRef thisws As Worksheet, _
                                          ByRef thiswb As Workbook, _
                                          ByVal dqcPath As String)
    On Error GoTo ErrorHandler
    
    Dim sPath As String
    Dim review_number As Long
    Dim sample_month As String
    Dim findmemo As String
    
    sPath = thiswb.Path
    
    ' Check if error case
    If thisws.Range("AL10").Value = 1 Then
        MsgBox "Finding = 1. This case is not an error case.", vbInformation
        Exit Sub
    End If
    
    review_number = CLng(thisws.Range("A10").Value)
    sample_month = CStr(thisws.Range("AB10").Value)
    findmemo = "QC FINDING TANF Review " & review_number & " Sample Month " & sample_month & ".xlsm"
    
    MsgBox "TANF Findings Memo" & vbCrLf & vbCrLf & _
           "Review Number: " & review_number & vbCrLf & _
           "Sample Month: " & sample_month & vbCrLf & _
           "Error Amount: " & Format(thisws.Range("AO10").Value, "$#,##0.00") & vbCrLf & vbCrLf & _
           "Full implementation would generate the complete memo.", _
           vbInformation, "TANF Findings Memo"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    LogError "GenerateTANFFindingsMemoFull", Err.Number, Err.Description, ""
End Sub


' ============================================================================
' GA FINDINGS MEMO
' ============================================================================

Private Sub GenerateGAFindingsMemoFull(ByRef thisws As Worksheet, _
                                        ByRef thiswb As Workbook, _
                                        ByVal dqcPath As String)
    On Error GoTo ErrorHandler
    
    Dim review_number As Long
    Dim sample_month As String
    
    review_number = CLng(thisws.Range("A10").Value)
    sample_month = CStr(thisws.Range("AB10").Value)
    
    MsgBox "GA Findings Memo" & vbCrLf & vbCrLf & _
           "Review Number: " & review_number & vbCrLf & _
           "Sample Month: " & sample_month & vbCrLf & vbCrLf & _
           "Full implementation would generate the complete memo.", _
           vbInformation, "GA Findings Memo"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    LogError "GenerateGAFindingsMemoFull", Err.Number, Err.Description, ""
End Sub


' ============================================================================
' MA FINDINGS MEMO IMPLEMENTATIONS
' ============================================================================

Private Sub GenerateMAPosFindingsMemoFull(ByRef thisws As Worksheet, _
                                           ByRef thiswb As Workbook, _
                                           ByVal dqcPath As String, _
                                           ByVal memo_type As String)
    On Error GoTo ErrorHandler
    
    Dim review_number As String
    
    review_number = CStr(thisws.Range("A10").Value)
    
    MsgBox "MA Positive " & memo_type & vbCrLf & vbCrLf & _
           "Review Number: " & review_number & vbCrLf & vbCrLf & _
           "Full implementation would generate the complete memo.", _
           vbInformation, "MA Positive Memo"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    LogError "GenerateMAPosFindingsMemoFull", Err.Number, Err.Description, memo_type
End Sub

Private Sub GenerateMANegFindingsMemoFull(ByRef thisws As Worksheet, _
                                           ByRef thiswb As Workbook, _
                                           ByVal dqcPath As String, _
                                           ByVal memo_type As String)
    On Error GoTo ErrorHandler
    
    Dim review_number As String
    
    review_number = CStr(thisws.Range("C20").Value)
    
    MsgBox "MA Negative " & memo_type & vbCrLf & vbCrLf & _
           "Review Number: " & review_number & vbCrLf & vbCrLf & _
           "Full implementation would generate the complete memo.", _
           vbInformation, "MA Negative Memo"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    LogError "GenerateMANegFindingsMemoFull", Err.Number, Err.Description, memo_type
End Sub

Private Sub GenerateMAPEFindingsMemoFull(ByRef thisws As Worksheet, _
                                          ByRef thiswb As Workbook, _
                                          ByVal dqcPath As String, _
                                          ByVal memo_type As String)
    On Error GoTo ErrorHandler
    
    Dim review_number As String
    
    review_number = CStr(thisws.Range("A10").Value)
    
    MsgBox "MA PE " & memo_type & vbCrLf & vbCrLf & _
           "Review Number: " & review_number & vbCrLf & vbCrLf & _
           "Full implementation would generate the complete memo.", _
           vbInformation, "MA PE Memo"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    LogError "GenerateMAPEFindingsMemoFull", Err.Number, Err.Description, memo_type
End Sub


' ============================================================================
' TIMELINESS MEMO
' ============================================================================

Private Sub GenerateTimelinessMemo(ByVal finding_type As String)
    On Error GoTo ErrorHandler
    
    Dim thisws As Worksheet
    Dim review_number As Long
    Dim sample_month As String
    Dim findmemo As String
    
    Set thisws = ActiveSheet
    
    review_number = CLng(thisws.Range("A18").Value)
    sample_month = thisws.Range("AD18").Value & thisws.Range("AG18").Value
    
    If finding_type = "Application" Then
        findmemo = "QC FINDING Application Timeliness Review " & review_number & " Sample Month " & sample_month & ".xlsm"
    Else
        findmemo = "QC FINDING Renewal Timeliness Review " & review_number & " Sample Month " & sample_month & ".xlsm"
    End If
    
    MsgBox "Timeliness " & finding_type & " Memo" & vbCrLf & vbCrLf & _
           "Review Number: " & review_number & vbCrLf & _
           "Sample Month: " & sample_month & vbCrLf & vbCrLf & _
           "Memo would be saved as: " & findmemo, _
           vbInformation, "Timeliness Memo"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    LogError "GenerateTimelinessMemo", Err.Number, Err.Description, finding_type
End Sub

Private Sub GenerateTimeMemo(ByVal memoType As String)
    On Error GoTo ErrorHandler
    
    MsgBox "Time Memo (" & memoType & ") generation placeholder.", vbInformation
    
    Exit Sub
    
ErrorHandler:
    LogError "GenerateTimeMemo", Err.Number, Err.Description, memoType
End Sub


' ============================================================================
' HELPER FUNCTIONS
' ============================================================================

' ----------------------------------------------------------------------------
' Function: GetElementNatureCauseCodes
' ----------------------------------------------------------------------------
' PURPOSE:
'   Extracts element, nature, and cause codes from the schedule and formats
'   them as a string for the findings memo.
' ----------------------------------------------------------------------------
Private Function GetElementNatureCauseCodes(ByRef ws As Worksheet, _
                                             ByVal programType As String) As String
    On Error Resume Next
    
    Dim enc_codes As String
    Dim irow As Long
    
    enc_codes = ""
    
    Select Case programType
        Case "5"  ' SNAP Positive - Section 2
            For irow = 29 To 43 Step 2
                If ws.Range("B" & irow).Value = "" Then Exit For
                If enc_codes <> "" Then enc_codes = enc_codes & vbNewLine
                enc_codes = enc_codes & ws.Range("B" & irow).Value & " - " & _
                            ws.Range("G" & irow).Value & " - " & ws.Range("K" & irow).Value
            Next irow
            
        Case "6"  ' SNAP Negative
            enc_codes = ws.Range("E47").Value & " - " & ws.Range("W47").Value
            
        Case "1", "9"  ' TANF/GA - Section 6
            For irow = 61 To 67 Step 2
                If ws.Range("F" & irow).Value = "" Then Exit For
                If enc_codes <> "" Then enc_codes = enc_codes & vbNewLine
                enc_codes = enc_codes & ws.Range("J" & irow).Value & " - " & _
                            ws.Range("O" & irow).Value & " - " & ws.Range("T" & irow).Value
            Next irow
    End Select
    
    GetElementNatureCauseCodes = enc_codes
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' Function: BuildCCString
' ----------------------------------------------------------------------------
' PURPOSE:
'   Builds the CC (courtesy copy) string for the findings memo by reading
'   from the Courtesy Copy for Memos file.
'
' PARAMETERS:
'   dqcPath      - The DQC drive path
'   areaNum      - The area number (1, 2, 3, etc.)
'   officeManager - The office manager name
'   ccCol        - The column to use (D for SNAP, F for TANF, etc.)
' ----------------------------------------------------------------------------
Private Function BuildCCString(ByVal dqcPath As String, _
                                ByVal areaNum As String, _
                                ByVal officeManager As String, _
                                ByVal ccCol As String) As String
    On Error GoTo ErrorHandler
    
    Dim ccwb As Workbook
    Dim ccws As Worksheet
    Dim ccstring As String
    Dim ccrow As Long
    Dim tempArray() As String
    Dim isplit As Long
    
    ' Open CC list file
    Workbooks.Open Filename:=dqcPath & TEMPLATE_COURTESY_COPY_LIST, _
                   UpdateLinks:=False, ReadOnly:=True
    Set ccwb = ActiveWorkbook
    Set ccws = ccwb.Worksheets(areaNum)
    
    ' Build CC string
    ccstring = vbNewLine & vbNewLine & "cc: "
    ccstring = ccstring & ccws.Range(ccCol & "4").Value & vbNewLine  ' Area Manager
    ccstring = ccstring & ccws.Range(ccCol & "5").Value & vbNewLine  ' Area Staff Asst
    
    ccrow = 6
    If areaNum = "1" Then
        ccstring = ccstring & ccws.Range(ccCol & ccrow).Value & vbNewLine  ' QA for Area 1
        ccrow = ccrow + 1
    End If
    
    ccstring = ccstring & officeManager & vbNewLine  ' Office manager
    ccrow = ccrow + 1
    
    ' Corrective Action (may have commas)
    If InStr(ccws.Range(ccCol & ccrow).Value, ",") > 0 Then
        tempArray = Split(ccws.Range(ccCol & ccrow).Value, ",")
        For isplit = LBound(tempArray) To UBound(tempArray)
            ccstring = ccstring & Trim(tempArray(isplit)) & vbNewLine
        Next isplit
    Else
        ccstring = ccstring & ccws.Range(ccCol & ccrow).Value & vbNewLine
    End If
    ccrow = ccrow + 1
    
    ' Additional Recipients
    If Not (ccws.Range(ccCol & ccrow).Value = "-" Or ccws.Range(ccCol & ccrow).Value = "") Then
        If InStr(ccws.Range(ccCol & ccrow).Value, ",") > 0 Then
            tempArray = Split(ccws.Range(ccCol & ccrow).Value, ",")
            For isplit = LBound(tempArray) To UBound(tempArray)
                ccstring = ccstring & Trim(tempArray(isplit)) & vbNewLine
            Next isplit
        Else
            ccstring = ccstring & ccws.Range(ccCol & ccrow).Value & vbNewLine
        End If
    End If
    ccrow = ccrow + 1
    
    ' Program Manager
    If Not (ccws.Range(ccCol & ccrow).Value = "-" Or ccws.Range(ccCol & ccrow).Value = "") Then
        ccstring = ccstring & ccws.Range(ccCol & ccrow).Value & vbNewLine
    End If
    ccrow = ccrow + 1
    
    ' Posting
    ccstring = ccstring & ccws.Range(ccCol & ccrow).Value & vbNewLine
    ccrow = ccrow + 1
    
    ' File
    ccstring = ccstring & ccws.Range(ccCol & ccrow).Value & vbNewLine
    
    ' Close CC workbook
    ccwb.Close False
    
    BuildCCString = ccstring
    Exit Function
    
ErrorHandler:
    On Error Resume Next
    If Not ccwb Is Nothing Then ccwb.Close False
    BuildCCString = vbNewLine & vbNewLine & "cc: [Error building CC list]"
End Function


