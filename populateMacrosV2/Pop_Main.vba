Attribute VB_Name = "Pop_Main"
' ============================================================================
' Pop_Main - Main Schedule Population Module
' ============================================================================
' WHAT THIS MODULE DOES:
'   This is the main orchestration module for creating new QC review schedules.
'   It takes data from BIS (Benefits Information System) delimited files and
'   populates schedule templates with case information.
'
' THE POPULATION PROCESS:
'   1. User selects the program type and review month
'   2. User selects the File of Records (list of cases to review)
'   3. User selects the BIS delimited file (case data)
'   4. System creates a new schedule workbook
'   5. For each case: extracts data from BIS file, populates schedule template
'   6. Schedule is saved to the examiner's folder on the network
'
' PROGRAMS SUPPORTED:
'   - SNAP Positive (FS Positive)
'   - SNAP Negative (FS Negative)
'   - TANF
'   - GA (uses TANF process with some differences)
'   - MA Positive (separate flow)
'   - MA Negative (separate flow)
'
' KEY CONCEPTS:
'   - BIS File: Excel file exported from BIS with case/individual data
'   - File of Records: List of cases selected for QC review this month
'   - Schedule Template: Blank review form to be filled with case data
'   - Examiner Folder: Network location where completed schedules are saved
'
' PERFORMANCE OPTIMIZATIONS (V2):
'   - Uses array-based I/O instead of cell-by-cell reads/writes
'   - Disables screen updating and events during processing
'   - Uses Common_Utils for shared functions (no duplication)
'   - Uses Config_Settings for all constants
'
' CHANGE LOG:
'   2026-01-02  Refactored from populate_mod - added Option Explicit,
'               array-based I/O, uses Common_Utils, V2 commenting style
' ============================================================================

Option Explicit

' ============================================================================
' MODULE-LEVEL VARIABLES
' ============================================================================
' NOTE: In the refactored version, we minimize module-level variables.
' The original had many globals which made the code hard to follow.
' We keep only what's truly needed across multiple subs.
' ============================================================================

Private m_BISWorkbook As Workbook       ' The BIS delimited file
Private m_ScheduleWorkbook As Workbook  ' The output schedule workbook
Private m_Program As String             ' Current program type
Private m_ReviewMonth As Long           ' Review month in YYYYMM format
Private m_DQCPath As String             ' Path to DQC folder

' ============================================================================
' PUBLIC VARIABLES - Shared with program-specific population modules
' ============================================================================
' These variables are accessed by Pop_SNAP_Positive, Pop_TANF, etc.
' They provide the workbook references needed for population.
' ============================================================================
Public wb As Workbook                   ' Schedule workbook (alias for m_ScheduleWorkbook)
Public wb_bis As Workbook               ' BIS workbook (alias for m_BISWorkbook)
Public program As String                ' Program type (alias for m_Program)


' ============================================================================
' ENTRY POINTS
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: showpopulate_form50
' ----------------------------------------------------------------------------
' PURPOSE:
'   Shows the population form (UF_PopulateMain) which allows the user to select
'   program type and review month before starting population.
'
' NOTE: Function name kept as "showpopulate_form50" for backward compatibility,
'       but now calls the renamed UF_PopulateMain form.
' ----------------------------------------------------------------------------
Public Sub showpopulate_form50()
    On Error Resume Next
    UF_PopulateMain.Show
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' Sub: showpopulate_form3
' ----------------------------------------------------------------------------
' PURPOSE:
'   Shows UserForm3 for additional population options.
'
' NOTE: UserForm3 was deprecated in V2 (drive picker - now auto-detected).
'       This function is kept for backward compatibility but does nothing.
' ----------------------------------------------------------------------------
Public Sub showpopulate_form3()
    On Error Resume Next
    ' UserForm3 deprecated - drive auto-detection now in Common_Utils
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' Sub: redisplayform3
' ----------------------------------------------------------------------------
' PURPOSE:
'   Redisplays UserForm3 (called after certain operations).
'
' NOTE: UserForm3 was deprecated in V2. This is a no-op for compatibility.
' ----------------------------------------------------------------------------
Public Sub redisplayform3()
    On Error Resume Next
    ' UserForm3 deprecated - drive auto-detection now in Common_Utils
    On Error GoTo 0
End Sub


' ============================================================================
' MAIN POPULATION FUNCTION
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: Review_Schedule
' ----------------------------------------------------------------------------
' PURPOSE:
'   Main entry point for creating review schedules. This is the primary
'   function that orchestrates the entire population process.
'
' STEPS:
'   1. Validate network path
'   2. Get program and month from user input (form)
'   3. Open File of Records and BIS file
'   4. Create output schedule workbook
'   5. Loop through cases and populate schedules
'   6. Save and close
'
' PRE-CONDITIONS:
'   - Network drive must be mapped
'   - User has selected program and month in the form
'   - BIS file must be in correct format
'
' POST-CONDITIONS:
'   - Schedule workbook created with populated data
'   - Individual schedule files created in examiner folders
' ----------------------------------------------------------------------------
Public Sub Review_Schedule()
    On Error GoTo ErrorHandler
    
    ' ========================================================================
    ' 1. PERFORMANCE OPTIMIZATION - Disable Excel overhead
    ' ========================================================================
    With Application
        .ScreenUpdating = False
        .DisplayStatusBar = True
        .EnableEvents = False
        .Calculation = xlCalculationManual
        .DisplayAlerts = False
    End With
    
    ' ========================================================================
    ' 2. VARIABLE DECLARATIONS
    ' ========================================================================
    Dim popWB As Workbook           ' This workbook (Populate.xlsm)
    Dim strCurDir As String         ' Current directory
    Dim inputFile As String         ' Path to File of Records
    Dim inputFileBIS As String      ' Path to BIS delimited file
    Dim monthStr As String          ' Month name for file naming
    Dim outputFileName As String    ' Output schedule filename
    
    Set popWB = ActiveWorkbook
    strCurDir = popWB.path & "\"
    
    ' ========================================================================
    ' 3. VALIDATE NETWORK PATH
    ' ========================================================================
    m_DQCPath = GetDQCDriveLetter()
    If m_DQCPath = "" Then
        MsgBox "Network Drive to Examiner Files is NOT correct" & vbCrLf & _
               "Contact Valerie or Nicole", vbCritical, "Network Error"
        GoTo CleanExit
    End If
    
    ' Validate schedules folder exists
    Dim schedulesPath As String
    schedulesPath = m_DQCPath & "Schedules by Examiner Number\"
    
    If Not PathExists(schedulesPath) Then
        MsgBox "Path to Examiner's Files: " & schedulesPath & " does NOT exist!!" & vbCrLf & _
               "Contact Valerie or Nicole", vbCritical, "Path Error"
        GoTo CleanExit
    End If
    
    ' ========================================================================
    ' 4. GET PROGRAM AND MONTH FROM USER INPUT
    ' ========================================================================
    ' These values come from the user form (cells on the populate sheet)
    m_Program = Cells(7, 23).Value     ' Column W, row 7
    m_ReviewMonth = Cells(7, 26).Value ' Column Z, row 7
    
    If m_Program = "" Then
        MsgBox "Please select a program type.", vbExclamation
        GoTo CleanExit
    End If
    
    ' ========================================================================
    ' 5. PROMPT FOR INPUT FILES
    ' ========================================================================
    
    ' Get File of Records
    inputFile = Application.GetOpenFilename( _
        "File of Records (*.xlsm), *.xlsm", , _
        "Select File of Records")
    
    If inputFile = "False" Then GoTo CleanExit
    
    ' Get BIS Delimited File (program-specific prompts)
    inputFileBIS = GetBISFileFromUser(m_Program)
    If inputFileBIS = "" Then GoTo CleanExit
    
    ' ========================================================================
    ' 6. CREATE OUTPUT WORKBOOK
    ' ========================================================================
    monthStr = Application.WorksheetFunction.Text(m_ReviewMonth, "MMMM YYYY")
    outputFileName = "Review Schedule for " & m_Program & " " & monthStr & ".xlsx"
    
    ' Create new workbook
    Workbooks.Add
    ActiveWorkbook.SaveAs Filename:=outputFileName, FileFormat:=51
    Set m_ScheduleWorkbook = ActiveWorkbook
    
    ' ========================================================================
    ' 7. OPEN INPUT FILES
    ' ========================================================================
    
    ' Open BIS file
    Workbooks.Open Filename:=inputFileBIS, UpdateLinks:=False
    Set m_BISWorkbook = ActiveWorkbook
    
    ' Open File of Records
    Workbooks.Open Filename:=inputFile, UpdateLinks:=False
    ' This workbook becomes active
    
    ' ========================================================================
    ' 8. PROCESS CASES
    ' ========================================================================
    ' Route to program-specific population
    Select Case m_Program
        Case "FS Positive", "FS Supplemental"
            Call PopulateSNAPPositive(popWB)
        Case "FS Negative"
            Call PopulateSNAPNegative(popWB)
        Case "TANF"
            Call PopulateTANF(popWB)
        Case "GA"
            Call PopulateGA(popWB)
        Case "MA Positive"
            Call PopulateMAPositive(popWB)
        Case "MA Negative"
            Call PopulateMANegative(popWB)
        Case Else
            MsgBox "Unknown program type: " & m_Program, vbExclamation
    End Select
    
    ' ========================================================================
    ' 9. CLEANUP
    ' ========================================================================
    MsgBox "Population complete for " & m_Program & "!", vbInformation, "Complete"
    
CleanExit:
    ' Restore Excel settings
    With Application
        .ScreenUpdating = True
        .DisplayStatusBar = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
        .DisplayAlerts = True
    End With
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Population Error"
    LogError "Review_Schedule", Err.Number, Err.Description, "Program: " & m_Program
    Resume CleanExit
End Sub


' ============================================================================
' PROGRAM-SPECIFIC POPULATION FUNCTIONS
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: PopulateSNAPPositive
' ----------------------------------------------------------------------------
' PURPOSE:
'   Handles the population process for SNAP Positive reviews.
'   Calls Pop_SNAP_Positive.PopulateSNAPPositiveDelimited after setting up
'   the global references needed by the population module.
'
' PARAMETERS:
'   popWB - The Populate.xlsm workbook (source of templates/macros)
'
' BIS FILE STRUCTURE:
'   - "Case" worksheet: One row per case with case-level data
'   - "Individual" worksheet: One row per household member
' ----------------------------------------------------------------------------
Private Sub PopulateSNAPPositive(ByRef popWB As Workbook)
    On Error GoTo ErrorHandler
    
    ' Set up global references for Pop_SNAP_Positive module
    ' These are accessed by the population module
    Set wb = m_ScheduleWorkbook
    Set wb_bis = m_BISWorkbook
    program = m_Program
    
    ' Call the SNAP Positive population module
    Call Pop_SNAP_Positive.PopulateSNAPPositiveDelimited
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in SNAP Positive population: " & Err.Description, vbCritical
    LogError "PopulateSNAPPositive", Err.Number, Err.Description, ""
End Sub

' ----------------------------------------------------------------------------
' Sub: PopulateSNAPNegative
' ----------------------------------------------------------------------------
' PURPOSE:
'   Handles the population process for SNAP Negative reviews.
'   Calls Pop_SNAP_Negative.PopulateSNAPNegativeDelimited.
' ----------------------------------------------------------------------------
Private Sub PopulateSNAPNegative(ByRef popWB As Workbook)
    On Error GoTo ErrorHandler
    
    ' Set up global references
    Set wb = m_ScheduleWorkbook
    Set wb_bis = m_BISWorkbook
    program = m_Program
    
    ' Call the SNAP Negative population module
    Call Pop_SNAP_Negative.PopulateSNAPNegativeDelimited
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in SNAP Negative population: " & Err.Description, vbCritical
    LogError "PopulateSNAPNegative", Err.Number, Err.Description, ""
End Sub

' ----------------------------------------------------------------------------
' Sub: PopulateTANF
' ----------------------------------------------------------------------------
' PURPOSE:
'   Handles the population process for TANF reviews.
'   Calls Pop_TANF.PopulateTANFDelimited.
' ----------------------------------------------------------------------------
Private Sub PopulateTANF(ByRef popWB As Workbook)
    On Error GoTo ErrorHandler
    
    ' Set up global references
    Set wb = m_ScheduleWorkbook
    Set wb_bis = m_BISWorkbook
    program = m_Program
    
    ' Call the TANF population module
    Call Pop_TANF.PopulateTANFDelimited
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in TANF population: " & Err.Description, vbCritical
    LogError "PopulateTANF", Err.Number, Err.Description, ""
End Sub

' ----------------------------------------------------------------------------
' Sub: PopulateGA
' ----------------------------------------------------------------------------
' PURPOSE:
'   Handles the population process for GA reviews.
'   GA uses TANF process with some differences (same BIS structure).
' ----------------------------------------------------------------------------
Private Sub PopulateGA(ByRef popWB As Workbook)
    On Error GoTo ErrorHandler
    
    ' Set up global references
    Set wb = m_ScheduleWorkbook
    Set wb_bis = m_BISWorkbook
    program = m_Program
    
    ' GA uses TANF population with GA-specific adjustments
    ' For now, route to TANF (can create Pop_GA if needed)
    Call Pop_TANF.PopulateTANFDelimited
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in GA population: " & Err.Description, vbCritical
    LogError "PopulateGA", Err.Number, Err.Description, ""
End Sub

' ----------------------------------------------------------------------------
' Sub: PopulateMAPositive
' ----------------------------------------------------------------------------
' PURPOSE:
'   Handles the population process for MA Positive reviews.
'   Calls Pop_MA.PopulateMADelimited.
' ----------------------------------------------------------------------------
Private Sub PopulateMAPositive(ByRef popWB As Workbook)
    On Error GoTo ErrorHandler
    
    ' Set up global references
    Set wb = m_ScheduleWorkbook
    Set wb_bis = m_BISWorkbook
    program = m_Program
    
    ' Call the MA population module
    Call Pop_MA.PopulateMADelimited
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in MA Positive population: " & Err.Description, vbCritical
    LogError "PopulateMAPositive", Err.Number, Err.Description, ""
End Sub

' ----------------------------------------------------------------------------
' Sub: PopulateMANegative
' ----------------------------------------------------------------------------
' PURPOSE:
'   Handles the population process for MA Negative reviews.
' ----------------------------------------------------------------------------
Private Sub PopulateMANegative(ByRef popWB As Workbook)
    On Error GoTo ErrorHandler
    
    ' Set up global references
    Set wb = m_ScheduleWorkbook
    Set wb_bis = m_BISWorkbook
    program = m_Program
    
    ' MA Negative uses same module as MA Positive
    Call Pop_MA.PopulateMADelimited
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in MA Negative population: " & Err.Description, vbCritical
    LogError "PopulateMANegative", Err.Number, Err.Description, ""
End Sub


' ============================================================================
' HELPER FUNCTIONS
' ============================================================================

' ----------------------------------------------------------------------------
' Function: GetBISFileFromUser
' ----------------------------------------------------------------------------
' PURPOSE:
'   Prompts the user to select the appropriate BIS delimited file based
'   on the program type.
'
' PARAMETERS:
'   programName - The program type string
'
' RETURNS:
'   String - Full path to selected file, or "" if cancelled
' ----------------------------------------------------------------------------
Private Function GetBISFileFromUser(ByVal programName As String) As String
    Dim filePath As Variant
    Dim dialogTitle As String
    Dim fileFilter As String
    
    Select Case programName
        Case "FS Positive", "FS Supplemental"
            dialogTitle = "Select BIS SNAP Positive Delimited Excel file"
            fileFilter = "BIS SNAP Positive File (*.xlsx), *.xlsx"
        Case "FS Negative"
            dialogTitle = "Select BIS SNAP Negative Delimited Excel file"
            fileFilter = "BIS SNAP Negative File (*.xlsx), *.xlsx"
        Case "TANF"
            dialogTitle = "Select BIS TANF Delimited Excel file"
            fileFilter = "BIS TANF File (*.xlsx), *.xlsx"
        Case "MA Positive"
            dialogTitle = "Select BIS MA Positive Delimited Excel file"
            fileFilter = "BIS MA Positive File (*.xlsx), *.xlsx"
        Case "MA Negative"
            dialogTitle = "Select BIS MA Negative Delimited Excel file"
            fileFilter = "BIS MA Negative File (*.xlsx), *.xlsx"
        Case Else
            GetBISFileFromUser = ""
            Exit Function
    End Select
    
    filePath = Application.GetOpenFilename(fileFilter, , dialogTitle)
    
    If filePath = False Then
        GetBISFileFromUser = ""
    Else
        GetBISFileFromUser = CStr(filePath)
    End If
End Function

' ----------------------------------------------------------------------------
' Sub: ReadBISDataToArrays
' ----------------------------------------------------------------------------
' PURPOSE:
'   Reads all data from the BIS file into memory arrays for fast processing.
'   This is the key V2 optimization - one read instead of thousands.
'
' PARAMETERS:
'   bisWB       - The BIS workbook
'   caseData    - Output array for case-level data
'   indivData   - Output array for individual-level data
'
' NOTE:
'   This is a template showing the array-based approach. The actual
'   implementation would read specific columns needed for population.
' ----------------------------------------------------------------------------
Private Sub ReadBISDataToArrays(ByRef bisWB As Workbook, _
                                 ByRef caseData As Variant, _
                                 ByRef indivData As Variant)
    On Error GoTo ErrorHandler
    
    Dim caseSheet As Worksheet
    Dim indivSheet As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    
    ' Get the Case worksheet
    On Error Resume Next
    Set caseSheet = bisWB.Worksheets("Case")
    If caseSheet Is Nothing Then
        Set caseSheet = bisWB.Worksheets(1)  ' Fallback to first sheet
    End If
    On Error GoTo ErrorHandler
    
    ' Find data extent
    lastRow = caseSheet.Cells(caseSheet.Rows.Count, 1).End(xlUp).Row
    lastCol = caseSheet.Cells(1, caseSheet.Columns.Count).End(xlToLeft).Column
    
    ' Read entire range into array (ONE read operation)
    ' This is MUCH faster than reading cell by cell
    caseData = caseSheet.Range(caseSheet.Cells(1, 1), _
                               caseSheet.Cells(lastRow, lastCol)).Value
    
    ' Get the Individual worksheet
    On Error Resume Next
    Set indivSheet = bisWB.Worksheets("Individual")
    If indivSheet Is Nothing Then
        Set indivSheet = bisWB.Worksheets(2)  ' Fallback to second sheet
    End If
    On Error GoTo ErrorHandler
    
    If Not indivSheet Is Nothing Then
        lastRow = indivSheet.Cells(indivSheet.Rows.Count, 1).End(xlUp).Row
        lastCol = indivSheet.Cells(1, indivSheet.Columns.Count).End(xlToLeft).Column
        
        indivData = indivSheet.Range(indivSheet.Cells(1, 1), _
                                     indivSheet.Cells(lastRow, lastCol)).Value
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error reading BIS data: " & Err.Description, vbCritical
    LogError "ReadBISDataToArrays", Err.Number, Err.Description, ""
End Sub

' ----------------------------------------------------------------------------
' Sub: WriteScheduleFromArray
' ----------------------------------------------------------------------------
' PURPOSE:
'   Writes data to a schedule worksheet from an array. This is the V2
'   optimization for output - one write instead of thousands.
'
' PARAMETERS:
'   ws        - Target worksheet
'   dataArray - 2D array of data to write
'   startRow  - Starting row on the worksheet
'   startCol  - Starting column on the worksheet
' ----------------------------------------------------------------------------
Private Sub WriteScheduleFromArray(ByRef ws As Worksheet, _
                                    ByRef dataArray As Variant, _
                                    ByVal startRow As Long, _
                                    ByVal startCol As Long)
    On Error GoTo ErrorHandler
    
    Dim numRows As Long
    Dim numCols As Long
    
    ' Get array dimensions
    numRows = UBound(dataArray, 1) - LBound(dataArray, 1) + 1
    numCols = UBound(dataArray, 2) - LBound(dataArray, 2) + 1
    
    ' Write entire array in ONE operation
    ws.Cells(startRow, startCol).Resize(numRows, numCols).Value = dataArray
    
    Exit Sub
    
ErrorHandler:
    LogError "WriteScheduleFromArray", Err.Number, Err.Description, ""
End Sub

' ----------------------------------------------------------------------------
' Function: BuildExaminerFolderPath
' ----------------------------------------------------------------------------
' PURPOSE:
'   Builds the full path to an examiner's program folder.
'
' PARAMETERS:
'   examinerName - The examiner's name
'   examinerNum  - The examiner's number
'   programName  - The program folder name
'   monthName    - The review month name
'   yearStr      - The 4-digit year
'
' RETURNS:
'   String - Full path to the folder
' ----------------------------------------------------------------------------
Private Function BuildExaminerFolderPath(ByVal examinerName As String, _
                                          ByVal examinerNum As String, _
                                          ByVal programName As String, _
                                          ByVal monthName As String, _
                                          ByVal yearStr As String) As String
    
    Dim basePath As String
    basePath = m_DQCPath & "Schedules by Examiner Number\"
    
    BuildExaminerFolderPath = basePath & _
                              examinerName & " - " & FormatExaminerNumber(examinerNum) & "\" & _
                              programName & "\" & _
                              "Review Month " & monthName & " " & yearStr & "\"
End Function


