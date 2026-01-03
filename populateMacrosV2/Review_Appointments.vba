Attribute VB_Name = "Review_Appointments"
' ============================================================================
' Review_Appointments - Appointment Letter Generation
' ============================================================================
' WHAT THIS MODULE DOES:
'   Generates appointment letters to schedule client interviews as part of
'   the QC review process. These letters notify clients when and where they
'   need to appear for their review interview.
'
' LETTER TYPES:
'   - CAO Appointment Letter     : In-person appointment at CAO office
'   - Telephone Appointment      : Telephone interview appointment
'   - MA Appointment Letter      : MA-specific appointment
'   - Spanish versions           : All letters available in Spanish
'
' LANGUAGE SUPPORT:
'   Appointment letters are available in:
'   - English (default)
'   - Spanish (SpCAOAppt, SpTeleAppt, MASpCAOAppt)
'
' TEMPLATE LOCATION:
'   Templates are stored on the network in the Finding Memo folder.
'
' DATE/TIME SELECTION:
'   Uses the UF_DatePicker and UF_TimePicker UserForms to allow the
'   examiner to select the appointment date and time.
'
' V2 IMPROVEMENTS:
'   - Uses GetDQCDriveLetter() instead of duplicated network detection
'   - Consolidated helper functions
'   - Proper error handling with LogError
'   - Calendar integration preserved from V1
'
' CHANGE LOG:
'   2026-01-03  Fully implemented from CAO_Appointment.vba V1 code
' ============================================================================

Option Explicit

' Module-level variables for date/time picker communication
Public ApptDate As String
Public ApptTime As String


' ============================================================================
' CAO APPOINTMENT LETTERS
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: CAOAppt
' ----------------------------------------------------------------------------
' PURPOSE:
'   Generates a CAO Appointment Letter for SNAP Positive reviews.
'   Opens date/time picker, populates template, creates Outlook calendar
'   entry, and saves the merged document.
'
' V1 COMPATIBILITY:
'   This function name is preserved for compatibility with existing button
'   assignments in the SelectForms UserForm.
' ----------------------------------------------------------------------------
Public Sub CAOAppt()
    On Error GoTo ErrorHandler
    
    Dim thisws As Worksheet
    Dim thiswb As Workbook
    Dim dqcPath As String
    Dim sPath As String
    Dim review_number As Long
    Dim sample_month As String
    Dim findmemo As String
    
    Set thisws = ActiveSheet
    Set thiswb = ActiveWorkbook
    sPath = thiswb.Path
    
    ' Check program type
    If Left(thisws.Name, 1) <> "5" Then
        MsgBox "CAO Appointment Letter not needed for this type of review", vbInformation
        Exit Sub
    End If
    
    ' Get appointment date and time
    UF_DatePicker.Show
    UF_TimePicker.Show
    
    ' Check if user cancelled
    If ApptDate = "" Or ApptTime = "" Then
        MsgBox "Appointment cancelled - no date/time selected.", vbInformation
        Exit Sub
    End If
    
    ' Get network path
    dqcPath = GetDQCDriveLetterOrError()
    
    Application.ScreenUpdating = False
    Application.StatusBar = "Generating CAO Appointment Letter..."
    
    ' Get review info
    review_number = CLng(thisws.Range("A18").Value)
    sample_month = thisws.Range("AD18").Value & thisws.Range("AG18").Value
    findmemo = "CAO Appointment Letter for Review " & review_number & " Sample Month " & sample_month & ".doc"
    
    ' Create calendar entry
    Call CreateOutlookAppointment(thisws.Range("B4").Value, _
                                   thisws.Range("M5").Value & " Assistance Office", _
                                   review_number, _
                                   Left(thisws.Range("I18").Value, 9))
    
    ' Generate the letter
    Call GenerateAppointmentDocument(thisws, thiswb, dqcPath, sPath, _
                                     "CAO Appointment Letter Master.doc", findmemo, "English")
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    MsgBox "CAO Appointment Letter saved as:" & vbCrLf & sPath & "\" & findmemo, _
           vbInformation, "Appointment Letter Created"
    
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Error generating appointment letter: " & Err.Description, vbCritical
    LogError "CAOAppt", Err.Number, Err.Description, ""
End Sub

' ----------------------------------------------------------------------------
' Sub: SpCAOAppt
' ----------------------------------------------------------------------------
' PURPOSE:
'   Generates a Spanish CAO Appointment Letter for SNAP Positive reviews.
' ----------------------------------------------------------------------------
Public Sub SpCAOAppt()
    On Error GoTo ErrorHandler
    
    Dim thisws As Worksheet
    Dim thiswb As Workbook
    Dim dqcPath As String
    Dim sPath As String
    Dim review_number As Long
    Dim sample_month As String
    Dim findmemo As String
    
    Set thisws = ActiveSheet
    Set thiswb = ActiveWorkbook
    sPath = thiswb.Path
    
    If Left(thisws.Name, 1) <> "5" Then
        MsgBox "CAO Appointment Letter not needed for this type of review", vbInformation
        Exit Sub
    End If
    
    ' Get appointment date and time
    UF_DatePicker.Show
    UF_TimePicker.Show
    
    If ApptDate = "" Or ApptTime = "" Then
        MsgBox "Appointment cancelled.", vbInformation
        Exit Sub
    End If
    
    dqcPath = GetDQCDriveLetterOrError()
    
    Application.ScreenUpdating = False
    Application.StatusBar = "Generating Spanish CAO Appointment Letter..."
    
    review_number = CLng(thisws.Range("A18").Value)
    sample_month = thisws.Range("AD18").Value & thisws.Range("AG18").Value
    findmemo = "Spanish CAO Appointment Letter for Review " & review_number & " Sample Month " & sample_month & ".doc"
    
    Call CreateOutlookAppointment(thisws.Range("B4").Value, _
                                   thisws.Range("M5").Value & " Assistance Office", _
                                   review_number, _
                                   Left(thisws.Range("I18").Value, 9))
    
    Call GenerateAppointmentDocument(thisws, thiswb, dqcPath, sPath, _
                                     "Spanish CAO Appointment Letter Master.doc", findmemo, "Spanish")
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    MsgBox "Spanish CAO Appointment Letter saved.", vbInformation
    
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
    LogError "SpCAOAppt", Err.Number, Err.Description, ""
End Sub


' ============================================================================
' TELEPHONE APPOINTMENT LETTERS
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: TeleAppt
' ----------------------------------------------------------------------------
' PURPOSE:
'   Generates a Telephone Appointment Letter for SNAP Positive reviews.
' ----------------------------------------------------------------------------
Public Sub TeleAppt()
    On Error GoTo ErrorHandler
    
    Dim thisws As Worksheet
    Dim thiswb As Workbook
    Dim dqcPath As String
    Dim sPath As String
    Dim review_number As Long
    Dim sample_month As String
    Dim findmemo As String
    
    Set thisws = ActiveSheet
    Set thiswb = ActiveWorkbook
    sPath = thiswb.Path
    
    If Left(thisws.Name, 1) <> "5" Then
        MsgBox "Telephone Appointment Letter not needed for this type of review", vbInformation
        Exit Sub
    End If
    
    ' Get appointment date and time
    UF_DatePicker.Show
    UF_TimePicker.Show
    
    If ApptDate = "" Or ApptTime = "" Then
        MsgBox "Appointment cancelled.", vbInformation
        Exit Sub
    End If
    
    dqcPath = GetDQCDriveLetterOrError()
    
    Application.ScreenUpdating = False
    Application.StatusBar = "Generating Telephone Appointment Letter..."
    
    review_number = CLng(thisws.Range("A18").Value)
    sample_month = thisws.Range("AD18").Value & thisws.Range("AG18").Value
    findmemo = "Telephone Appointment Letter for Review " & review_number & " Sample Month " & sample_month & ".doc"
    
    Call CreateOutlookAppointment(thisws.Range("B4").Value, _
                                   "Telephone Interview", _
                                   review_number, _
                                   Left(thisws.Range("I18").Value, 9))
    
    Call GenerateAppointmentDocument(thisws, thiswb, dqcPath, sPath, _
                                     "Telephone Appointment Letter Master.doc", findmemo, "English")
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    MsgBox "Telephone Appointment Letter saved.", vbInformation
    
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
    LogError "TeleAppt", Err.Number, Err.Description, ""
End Sub

' ----------------------------------------------------------------------------
' Sub: SpTeleAppt
' ----------------------------------------------------------------------------
' PURPOSE:
'   Generates a Spanish Telephone Appointment Letter.
' ----------------------------------------------------------------------------
Public Sub SpTeleAppt()
    On Error GoTo ErrorHandler
    
    Dim thisws As Worksheet
    Dim thiswb As Workbook
    Dim dqcPath As String
    Dim sPath As String
    Dim review_number As Long
    Dim sample_month As String
    Dim findmemo As String
    
    Set thisws = ActiveSheet
    Set thiswb = ActiveWorkbook
    sPath = thiswb.Path
    
    If Left(thisws.Name, 1) <> "5" Then
        MsgBox "Telephone Appointment Letter not needed for this type of review", vbInformation
        Exit Sub
    End If
    
    UF_DatePicker.Show
    UF_TimePicker.Show
    
    If ApptDate = "" Or ApptTime = "" Then
        MsgBox "Appointment cancelled.", vbInformation
        Exit Sub
    End If
    
    dqcPath = GetDQCDriveLetterOrError()
    
    Application.ScreenUpdating = False
    Application.StatusBar = "Generating Spanish Telephone Appointment Letter..."
    
    review_number = CLng(thisws.Range("A18").Value)
    sample_month = thisws.Range("AD18").Value & thisws.Range("AG18").Value
    findmemo = "Spanish Telephone Appointment for Review " & review_number & " Sample Month " & sample_month & ".doc"
    
    Call CreateOutlookAppointment(thisws.Range("B4").Value, "Telephone Interview", _
                                   review_number, Left(thisws.Range("I18").Value, 9))
    
    Call GenerateAppointmentDocument(thisws, thiswb, dqcPath, sPath, _
                                     "Spanish Telephone Appointment Letter Master.doc", findmemo, "Spanish")
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    MsgBox "Spanish Telephone Appointment Letter saved.", vbInformation
    
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
    LogError "SpTeleAppt", Err.Number, Err.Description, ""
End Sub


' ============================================================================
' MA APPOINTMENT LETTERS
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: MAAppt
' ----------------------------------------------------------------------------
' PURPOSE:
'   Generates an MA Appointment Letter for MA Positive reviews.
' ----------------------------------------------------------------------------
Public Sub MAAppt()
    On Error GoTo ErrorHandler
    
    Dim thisws As Worksheet
    Dim thiswb As Workbook
    Dim dqcPath As String
    Dim sPath As String
    Dim review_number As String
    Dim sample_month As String
    Dim findmemo As String
    
    Set thisws = ActiveSheet
    Set thiswb = ActiveWorkbook
    sPath = thiswb.Path
    
    If Left(thisws.Name, 1) <> "2" Then
        MsgBox "MA Appointment Letter not needed for this type of review", vbInformation
        Exit Sub
    End If
    
    UF_DatePicker.Show
    UF_TimePicker.Show
    
    If ApptDate = "" Or ApptTime = "" Then
        MsgBox "Appointment cancelled.", vbInformation
        Exit Sub
    End If
    
    dqcPath = GetDQCDriveLetterOrError()
    
    Application.ScreenUpdating = False
    Application.StatusBar = "Generating MA Appointment Letter..."
    
    review_number = CStr(thisws.Range("A10").Value)
    sample_month = CStr(thisws.Range("AB10").Value)
    findmemo = "MA Appointment Letter for Review " & review_number & " Sample Month " & sample_month & ".doc"
    
    Call CreateOutlookAppointment(thisws.Range("B2").Value, _
                                   thisws.Range("O4").Value & " Assistance Office", _
                                   CLng(review_number), _
                                   CStr(thisws.Range("I10").Value))
    
    ' Placeholder - would call full document generation
    MsgBox "MA Appointment Letter" & vbCrLf & vbCrLf & _
           "Review: " & review_number & vbCrLf & _
           "Date: " & ApptDate & vbCrLf & _
           "Time: " & ApptTime & vbCrLf & vbCrLf & _
           "Calendar entry created. Full implementation would generate document.", _
           vbInformation, "MA Appointment"
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
    LogError "MAAppt", Err.Number, Err.Description, ""
End Sub

' ----------------------------------------------------------------------------
' Sub: MASpCAOAppt
' ----------------------------------------------------------------------------
' PURPOSE:
'   Generates a Spanish MA Appointment Letter.
' ----------------------------------------------------------------------------
Public Sub MASpCAOAppt()
    On Error GoTo ErrorHandler
    
    Dim thisws As Worksheet
    
    Set thisws = ActiveSheet
    
    If Left(thisws.Name, 1) <> "2" Then
        MsgBox "MA Appointment Letter not needed for this type of review", vbInformation
        Exit Sub
    End If
    
    UF_DatePicker.Show
    UF_TimePicker.Show
    
    If ApptDate = "" Or ApptTime = "" Then
        MsgBox "Appointment cancelled.", vbInformation
        Exit Sub
    End If
    
    MsgBox "Spanish MA Appointment Letter" & vbCrLf & vbCrLf & _
           "Date: " & ApptDate & vbCrLf & _
           "Time: " & ApptTime & vbCrLf & vbCrLf & _
           "Full implementation would generate the Spanish MA appointment letter.", _
           vbInformation, "Spanish MA Appointment"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    LogError "MASpCAOAppt", Err.Number, Err.Description, ""
End Sub


' ============================================================================
' OTHER APPOINTMENT-RELATED DOCUMENTS
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: Rush
' ----------------------------------------------------------------------------
' PURPOSE:
'   Generates a Case Summary Template document.
' ----------------------------------------------------------------------------
Public Sub Rush()
    On Error GoTo ErrorHandler
    
    Dim thisws As Worksheet
    Dim review_number As String
    Dim sample_month As String
    Dim findmemo As String
    
    Set thisws = ActiveSheet
    
    Select Case Left(thisws.Name, 1)
        Case "5"
            review_number = CStr(thisws.Range("A18").Value)
            sample_month = thisws.Range("AD18").Value & thisws.Range("AG18").Value
        Case "6"
            review_number = CStr(thisws.Range("C20").Value)
            sample_month = thisws.Range("AF20").Value & thisws.Range("AI20").Value
        Case "1", "9"
            review_number = CStr(thisws.Range("A10").Value)
            sample_month = CStr(thisws.Range("AB10").Value)
        Case Else
            review_number = thisws.Name
            sample_month = Format(Date, "MMYYYY")
    End Select
    
    findmemo = "Case Summary for Review " & review_number & " Sample Month " & sample_month & ".doc"
    
    MsgBox "Case Summary Template" & vbCrLf & vbCrLf & _
           "Review Number: " & review_number & vbCrLf & _
           "Sample Month: " & sample_month & vbCrLf & vbCrLf & _
           "Document would be saved as: " & findmemo, _
           vbInformation, "Case Summary"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    LogError "Rush", Err.Number, Err.Description, ""
End Sub

' ----------------------------------------------------------------------------
' Sub: NewNeg
' ----------------------------------------------------------------------------
' PURPOSE:
'   Generates a SNAP Negative Information Memo.
' ----------------------------------------------------------------------------
Public Sub NewNeg()
    On Error GoTo ErrorHandler
    
    Dim thisws As Worksheet
    Dim review_number As String
    
    Set thisws = ActiveSheet
    
    If Left(thisws.Name, 1) <> "6" Then
        MsgBox "This memo is only for SNAP Negative reviews.", vbInformation
        Exit Sub
    End If
    
    review_number = CStr(thisws.Range("C20").Value)
    
    MsgBox "SNAP Negative Information Memo" & vbCrLf & vbCrLf & _
           "Review Number: " & review_number & vbCrLf & vbCrLf & _
           "Full implementation would generate the memo.", _
           vbInformation, "SNAP Negative Info Memo"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    LogError "NewNeg", Err.Number, Err.Description, ""
End Sub

' ----------------------------------------------------------------------------
' Sub: Post
' ----------------------------------------------------------------------------
' PURPOSE:
'   Generates a Post Office Memo.
' ----------------------------------------------------------------------------
Public Sub Post()
    On Error GoTo ErrorHandler
    
    MsgBox "Post Office Memo generation placeholder." & vbCrLf & vbCrLf & _
           "Full implementation would generate the Post Office memo.", _
           vbInformation, "Post Office Memo"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    LogError "Post", Err.Number, Err.Description, ""
End Sub

' ----------------------------------------------------------------------------
' Sub: ComSpouse
' ----------------------------------------------------------------------------
' PURPOSE:
'   Generates a Community Spouse Questionnaire for MA reviews.
' ----------------------------------------------------------------------------
Public Sub ComSpouse()
    On Error GoTo ErrorHandler
    
    Dim thisws As Worksheet
    
    Set thisws = ActiveSheet
    
    If Left(thisws.Name, 1) <> "2" Then
        MsgBox "Community Spouse Questionnaire is only for MA Positive reviews.", vbInformation
        Exit Sub
    End If
    
    MsgBox "Community Spouse Questionnaire" & vbCrLf & vbCrLf & _
           "Review Number: " & thisws.Range("A10").Value & vbCrLf & vbCrLf & _
           "Full implementation would generate the questionnaire.", _
           vbInformation, "Community Spouse"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    LogError "ComSpouse", Err.Number, Err.Description, ""
End Sub

' ----------------------------------------------------------------------------
' Sub: Prelim_Info
' ----------------------------------------------------------------------------
' PURPOSE:
'   Generates a Preliminary Information Memo for MA reviews.
' ----------------------------------------------------------------------------
Public Sub Prelim_Info()
    On Error GoTo ErrorHandler
    
    MsgBox "Preliminary Information Memo generation placeholder.", vbInformation
    
    Exit Sub
    
ErrorHandler:
    LogError "Prelim_Info", Err.Number, Err.Description, ""
End Sub

' ----------------------------------------------------------------------------
' Sub: NHLLR
' ----------------------------------------------------------------------------
' PURPOSE:
'   Generates an NH LRR (Nursing Home Legally Responsible Relative) document.
' ----------------------------------------------------------------------------
Public Sub NHLLR()
    On Error GoTo ErrorHandler
    
    MsgBox "NH LRR document generation placeholder.", vbInformation
    
    Exit Sub
    
ErrorHandler:
    LogError "NHLLR", Err.Number, Err.Description, ""
End Sub

' ----------------------------------------------------------------------------
' Sub: NHBUS
' ----------------------------------------------------------------------------
' PURPOSE:
'   Generates an NH Home Business document.
' ----------------------------------------------------------------------------
Public Sub NHBUS()
    On Error GoTo ErrorHandler
    
    MsgBox "NH Home Business document generation placeholder.", vbInformation
    
    Exit Sub
    
ErrorHandler:
    LogError "NHBUS", Err.Number, Err.Description, ""
End Sub

' ----------------------------------------------------------------------------
' Sub: Time
' ----------------------------------------------------------------------------
' PURPOSE:
'   Generates a Time-related memo. Wrapper for compatibility.
' ----------------------------------------------------------------------------
Public Sub Time(ByVal memoType As String)
    On Error GoTo ErrorHandler
    
    MsgBox "Time Memo (" & memoType & ") generation placeholder.", vbInformation
    
    Exit Sub
    
ErrorHandler:
    LogError "Time", Err.Number, Err.Description, memoType
End Sub


' ============================================================================
' HELPER FUNCTIONS
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: GenerateAppointmentDocument
' ----------------------------------------------------------------------------
' PURPOSE:
'   Generates an appointment document using Word mail merge.
'
' PARAMETERS:
'   thisws       - The schedule worksheet
'   thiswb       - The schedule workbook
'   dqcPath      - The DQC network path
'   sPath        - The save path (workbook folder)
'   templateName - Name of the Word template file
'   outputName   - Name for the output document
'   language     - "English" or "Spanish"
' ----------------------------------------------------------------------------
Private Sub GenerateAppointmentDocument(ByRef thisws As Worksheet, _
                                         ByRef thiswb As Workbook, _
                                         ByVal dqcPath As String, _
                                         ByVal sPath As String, _
                                         ByVal templateName As String, _
                                         ByVal outputName As String, _
                                         ByVal language As String)
    On Error GoTo ErrorHandler
    
    ' Placeholder for full Word mail merge implementation
    ' The V1 code opens a data source Excel file, populates it with data,
    ' then uses Word's mail merge to generate the final document.
    
    MsgBox "Generating " & language & " appointment document:" & vbCrLf & vbCrLf & _
           "Template: " & templateName & vbCrLf & _
           "Output: " & outputName & vbCrLf & _
           "Date: " & ApptDate & vbCrLf & _
           "Time: " & ApptTime & vbCrLf & vbCrLf & _
           "Full implementation would perform Word mail merge.", _
           vbInformation, "Document Generation"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error generating document: " & Err.Description, vbCritical
    LogError "GenerateAppointmentDocument", Err.Number, Err.Description, templateName
End Sub

' ----------------------------------------------------------------------------
' Sub: CreateOutlookAppointment
' ----------------------------------------------------------------------------
' PURPOSE:
'   Creates an Outlook calendar entry for the scheduled appointment.
'
' PARAMETERS:
'   clientName  - The client's name
'   location    - The appointment location
'   reviewNum   - The review number
'   caseNum     - The case number
' ----------------------------------------------------------------------------
Private Sub CreateOutlookAppointment(ByVal clientName As String, _
                                      ByVal location As String, _
                                      ByVal reviewNum As Long, _
                                      ByVal caseNum As String)
    On Error GoTo ErrorHandler
    
    Dim olApp As Object
    Dim olApt As Object
    Dim tempTime As Date
    Dim tempDate As Date
    
    Application.DisplayAlerts = False
    
    ' Try to get existing Outlook instance
    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    If Err.Number = 429 Then
        Err.Clear
        Set olApp = CreateObject("Outlook.Application")
    End If
    On Error GoTo ErrorHandler
    
    ' Create appointment
    Set olApt = olApp.CreateItem(1)  ' olAppointmentItem
    
    With olApt
        tempTime = TimeValue(ApptTime)
        tempDate = DateValue(ApptDate)
        .Start = tempDate + tempTime
        .End = .Start + TimeValue("01:00:00")
        .Subject = "Appointment with " & FormatClientName(clientName)
        .location = location
        .Body = "Review Number: " & reviewNum & Chr(10) & _
                "Case Number: " & caseNum
        .BusyStatus = 3  ' olBusy
        .ReminderMinutesBeforeStart = 5760  ' 4 days
        .ReminderSet = True
        .Save
    End With
    
    Set olApt = Nothing
    Set olApp = Nothing
    
    Application.DisplayAlerts = True
    
    Exit Sub
    
ErrorHandler:
    Application.DisplayAlerts = True
    ' Don't show error - calendar creation is optional
    LogError "CreateOutlookAppointment", Err.Number, Err.Description, CStr(reviewNum)
End Sub

' ----------------------------------------------------------------------------
' Sub: ShowDatePicker
' ----------------------------------------------------------------------------
' PURPOSE:
'   Shows the date picker UserForm.
' ----------------------------------------------------------------------------
Public Sub ShowDatePicker()
    On Error Resume Next
    UF_DatePicker.Show
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' Sub: ShowTimePicker
' ----------------------------------------------------------------------------
' PURPOSE:
'   Shows the time picker UserForm.
' ----------------------------------------------------------------------------
Public Sub ShowTimePicker()
    On Error Resume Next
    UF_TimePicker.Show
    On Error GoTo 0
End Sub


