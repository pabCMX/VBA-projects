Attribute VB_Name = "Pop_Repopulate"
' ============================================================================
' Pop_Repopulate - Module Update Tool for QC Review Schedules
' ============================================================================
' WHAT THIS MODULE DOES:
'   This module updates existing QC review schedules with the latest VBA code
'   (modules and UserForms). When bugs are fixed or features are added to the
'   macros, examiners' existing schedules still have the OLD code. This tool
'   opens those schedules and copies the new modules into them.
'
' WHY "REPOPULATE"?
'   - "Populate" creates NEW schedules from BIS files
'   - "Repopulate" UPDATES existing schedules with new code
'   It's like a software update, but for VBA macros inside workbooks.
'
' WHEN TO USE:
'   - After fixing a bug in one of the Review_* modules
'   - After adding a new feature examiners need
'   - When updating UserForms with new options
'   The administrator fills in the "repop" sheet with review numbers,
'   examiner numbers, and sample months, then runs the macro.
'
' HOW IT WORKS:
'   1. Reads a list of reviews from the "repop" worksheet
'   2. For each review:
'      a. Looks up the examiner name from the number
'      b. Determines the program type from the review number
'      c. Builds the file path based on examiner/program/month
'      d. Searches for the review workbook
'      e. Opens it and copies the appropriate modules/forms
'      f. Saves and closes the workbook
'
' INPUT SHEET LAYOUT (worksheet "repop"):
'   Column E: Review Number (e.g., 50012345)
'   Column F: Sample Month (YYYYMM format, e.g., 202401)
'   Column G: Examiner Number (e.g., 7)
'   Column K: Examiner Name (for lookup)
'   Column L: Examiner Number (for lookup)
'
' MODULES COPIED BY PROGRAM:
'   Each program needs different modules. See GetModuleListForProgram()
'   and GetFormListForProgram() for the complete lists.
'
' PERFORMANCE NOTES:
'   - Uses batch array reads to get review data (faster than cell-by-cell)
'   - Uses Dictionary for examiner lookup (O(1) instead of O(n))
'   - Replaced clFileSearchModule with native Dir() function
'   - Screen updating is disabled during processing
'
' REQUIREMENTS:
'   - This workbook must contain all the modules/forms to be copied
'   - Network drive must be mapped to the DQC folder
'   - User needs write access to examiner folders
'
' CHANGE LOG:
'   2026-01-02  Refactored from repopulate_mod - added Option Explicit,
'               array-based I/O, Dictionary lookups, V2 commenting style
' ============================================================================

Option Explicit

' ============================================================================
' MODULE-LEVEL CONSTANTS
' ============================================================================
' PURPOSE:
'   These constants define the worksheet and column layout. If the "repop"
'   sheet layout changes, update these constants.
' ============================================================================

Private Const REPOP_SHEET_NAME As String = "repop"
Private Const COL_REVIEW_NUM As Long = 5      ' Column E
Private Const COL_SAMPLE_MONTH As Long = 6    ' Column F
Private Const COL_EXAMINER_NUM As Long = 7    ' Column G
Private Const COL_EXAMINER_NAME As Long = 11  ' Column K (for lookup)
Private Const COL_EXAMINER_LOOKUP As Long = 12 ' Column L (for lookup)


' ============================================================================
' MAIN ENTRY POINT
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: RepopulateModulesAndForms
' ----------------------------------------------------------------------------
' PURPOSE:
'   Main entry point for the repopulation process. This is the sub that gets
'   attached to a button on the "repop" worksheet.
'
' STEPS:
'   1. Initialize - validate network, build examiner lookup
'   2. Loop through reviews - find and open each workbook
'   3. Copy modules/forms - based on program type
'   4. Cleanup - close workbooks, restore Excel settings
'
' ERROR HANDLING:
'   Uses On Error GoTo to catch problems. If something fails for one
'   review, it logs the error and continues with the next review.
' ----------------------------------------------------------------------------
Public Sub RepopulateModulesAndForms()
    On Error GoTo ErrorHandler
    
    ' ========================================================================
    ' 1. PERFORMANCE OPTIMIZATION - Disable Excel overhead
    ' ========================================================================
    ' These settings make the macro run MUCH faster by preventing Excel from
    ' constantly redrawing the screen and recalculating formulas.
    With Application
        .ScreenUpdating = False        ' Don't redraw screen after each operation
        .DisplayStatusBar = True       ' But show progress in status bar
        .EnableEvents = False          ' Don't trigger other macros
        .Calculation = xlCalculationManual  ' Don't recalculate formulas
    End With
    
    ' ========================================================================
    ' 2. DECLARE VARIABLES
    ' ========================================================================
    ' Note: We use Long instead of Integer throughout. Long is actually
    ' faster on modern (32-bit+) computers, AND it won't overflow if you
    ' have more than 32,767 rows.
    
    Dim dqcPath As String           ' Path to DQC folder (e.g., "E:\DQC\")
    Dim pathDir As String           ' Path to "Schedules by Examiner Number"
    Dim examinerLookup As Object    ' Dictionary: examNum -> examName
    Dim thisWB As Workbook          ' This workbook (contains modules to copy)
    Dim repopSheet As Worksheet     ' The "repop" worksheet with review list
    Dim scheduleWB As Workbook      ' The review workbook being updated
    
    Dim reviewData As Variant       ' Array of review numbers, months, examiners
    Dim examData As Variant         ' Array of examiner names and numbers
    Dim maxRow As Long              ' Last row of review data
    Dim maxRowEx As Long            ' Last row of examiner data
    
    Dim i As Long                   ' Loop counter
    Dim reviewNum As String         ' Current review number
    Dim monthStr As String          ' Current sample month (YYYYMM)
    Dim examNum As String           ' Current examiner number
    Dim examName As String          ' Current examiner name
    Dim examNumFormatted As String  ' Examiner number without leading zero
    Dim programType As ProgramType  ' Program enum (TANF, SNAP, etc.)
    Dim programFolder As String     ' Program folder name
    Dim monthName As String         ' Full month name (January, etc.)
    Dim yearStr As String           ' 4-digit year
    Dim searchPath As String        ' Path to search for workbook
    Dim foundFile As String         ' Name of found workbook file
    
    Dim processedCount As Long      ' Count of successfully processed reviews
    Dim errorCount As Long          ' Count of reviews that failed
    Dim startTime As Double         ' For tracking total time
    
    startTime = Timer
    processedCount = 0
    errorCount = 0
    
    ' ========================================================================
    ' 3. VALIDATE NETWORK PATH
    ' ========================================================================
    ' First thing we do is check if the network drive is mapped correctly.
    ' If not, there's no point continuing.
    
    dqcPath = GetDQCDriveLetter()
    If dqcPath = "" Then
        MsgBox "Network Drive to Examiner Files is NOT correct" & vbCrLf & _
               "Contact Valerie or Nicole", vbCritical, "Network Error"
        GoTo CleanExit
    End If
    
    pathDir = dqcPath & "Schedules by Examiner Number\"
    
    If Not PathExists(pathDir) Then
        MsgBox "Path to Examiner's Files: " & pathDir & " does NOT exist!!" & vbCrLf & _
               "Contact Valerie or Nicole", vbCritical, "Path Error"
        GoTo CleanExit
    End If
    
    ' ========================================================================
    ' 4. LOAD INPUT DATA INTO ARRAYS (One read operation)
    ' ========================================================================
    ' PERFORMANCE: Reading cells one-by-one is SLOW (each call crosses the
    ' VBA-Excel COM boundary). Reading the entire range at once is fast.
    ' After this, we access data via the arrays - no more Excel calls.
    
    Set thisWB = ActiveWorkbook
    Set repopSheet = thisWB.Worksheets(REPOP_SHEET_NAME)
    
    ' Find last row of data (using End(xlUp) from bottom is more reliable than xlDown)
    maxRow = repopSheet.Cells(repopSheet.Rows.Count, COL_REVIEW_NUM).End(xlUp).Row
    maxRowEx = repopSheet.Cells(repopSheet.Rows.Count, COL_EXAMINER_LOOKUP).End(xlUp).Row
    
    ' Read all review data in one operation
    ' We read columns E:G (review num, sample month, examiner num)
    reviewData = repopSheet.Range(repopSheet.Cells(1, COL_REVIEW_NUM), _
                                   repopSheet.Cells(maxRow, COL_EXAMINER_NUM)).Value
    
    ' Read all examiner lookup data in one operation
    ' We read columns K:L (examiner name, examiner number)
    examData = repopSheet.Range(repopSheet.Cells(1, COL_EXAMINER_NAME), _
                                 repopSheet.Cells(maxRowEx, COL_EXAMINER_LOOKUP)).Value
    
    ' ========================================================================
    ' 5. BUILD EXAMINER LOOKUP DICTIONARY
    ' ========================================================================
    ' PERFORMANCE: Looking up examiner names in a loop is O(n) - slow for
    ' many reviews. A Dictionary gives O(1) lookup - constant time.
    
    Set examinerLookup = CreateObject("Scripting.Dictionary")
    Call BuildExaminerDictionary(examData, examinerLookup)
    
    ' ========================================================================
    ' 6. MAIN PROCESSING LOOP
    ' ========================================================================
    ' For each review in the list:
    '   - Find the examiner
    '   - Determine the program
    '   - Build the file path
    '   - Find and open the workbook
    '   - Copy modules and forms
    '   - Save and close
    
    For i = 2 To maxRow  ' Start at row 2 to skip header
        
        ' Update status bar with progress
        Application.StatusBar = "Processing " & (i - 1) & "/" & (maxRow - 1) & _
                               " (" & Format((i - 1) / (maxRow - 1), "0%") & ")"
        DoEvents  ' Let Excel breathe (prevents "Not Responding")
        
        ' -----------------------------------------------------------------
        ' A. Get review number and clean it up
        ' -----------------------------------------------------------------
        reviewNum = Trim(CStr(reviewData(i, 1)))
        If reviewNum = "" Then GoTo NextReview
        
        ' Remove leading zero if present (e.g., "050012345" -> "50012345")
        reviewNum = RemoveLeadingZeros(reviewNum)
        
        ' -----------------------------------------------------------------
        ' B. Look up examiner name from Dictionary
        ' -----------------------------------------------------------------
        examNum = Trim(CStr(reviewData(i, 3)))
        
        If Not examinerLookup.Exists(examNum) Then
            MsgBox "No Examiner Name found for Review " & reviewNum & _
                   " and Examiner Number " & examNum & "." & vbCrLf & _
                   "Please check Review Number.", vbExclamation, "Examiner Not Found"
            errorCount = errorCount + 1
            GoTo NextReview
        End If
        
        examName = examinerLookup(examNum)
        examNumFormatted = FormatExaminerNumber(examNum)
        
        ' -----------------------------------------------------------------
        ' C. Determine program type from review number
        ' -----------------------------------------------------------------
        programType = GetProgramFromReviewNumber(reviewNum)
        
        If programType = PROG_UNKNOWN Then
            MsgBox "Review Number " & reviewNum & " is not a known QC Review number", _
                   vbExclamation, "Unknown Program"
            errorCount = errorCount + 1
            GoTo NextReview
        End If
        
        ' Get folder name for program (e.g., "SNAP Positive", "TANF")
        ' NOTE: The old code used "FS Positive" but folders may be "SNAP Positive"
        ' Adjust GetProgramFolderNameForRepop() if folder names differ
        programFolder = GetProgramFolderNameForRepop(programType)
        
        ' -----------------------------------------------------------------
        ' D. Parse sample month
        ' -----------------------------------------------------------------
        monthStr = Trim(CStr(reviewData(i, 2)))
        If Len(monthStr) < 6 Then
            errorCount = errorCount + 1
            GoTo NextReview
        End If
        
        yearStr = Left(monthStr, 4)
        monthName = MonthNumberToName(Right(monthStr, 2))
        
        If monthName = "" Then
            errorCount = errorCount + 1
            GoTo NextReview
        End If
        
        ' -----------------------------------------------------------------
        ' E. Build search path and find workbook
        ' -----------------------------------------------------------------
        ' The file path structure is:
        ' {DQC}\Schedules by Examiner Number\{Name} - {#}\{Program}\
        '     Review Month {Month} {Year}\{case folder}\Review Number...
        '
        ' We search the program folder for the file since the exact
        ' review month folder structure can vary.
        
        searchPath = pathDir & examName & " - " & examNumFormatted & "\" & _
                     programFolder & "\"
        
        If Not PathExists(searchPath) Then
            errorCount = errorCount + 1
            GoTo NextReview
        End If
        
        ' Find the workbook file (searches subdirectories)
        foundFile = FindReviewWorkbook(searchPath, reviewNum, monthStr)
        
        If foundFile = "" Then
            ' File not found - silently continue (this is expected for some reviews)
            GoTo NextReview
        End If
        
        ' -----------------------------------------------------------------
        ' F. Open workbook and copy modules/forms
        ' -----------------------------------------------------------------
        On Error Resume Next
        Set scheduleWB = Workbooks.Open(Filename:=foundFile, UpdateLinks:=False, ReadOnly:=False)
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo ErrorHandler
            errorCount = errorCount + 1
            GoTo NextReview
        End If
        On Error GoTo ErrorHandler
        
        ' Copy appropriate modules and forms for this program
        Call CopyModulesForProgram(thisWB, scheduleWB, programType)
        Call CopyFormsForProgram(thisWB, scheduleWB, programType)
        
        ' Save and close
        scheduleWB.Close SaveChanges:=True
        Set scheduleWB = Nothing
        
        processedCount = processedCount + 1
        
NextReview:
    Next i
    
    ' ========================================================================
    ' 7. SHOW COMPLETION MESSAGE
    ' ========================================================================
    Dim elapsedTime As Double
    elapsedTime = Timer - startTime
    
    MsgBox "Repopulation Complete!" & vbCrLf & vbCrLf & _
           "Reviews Processed: " & processedCount & vbCrLf & _
           "Errors/Skipped: " & errorCount & vbCrLf & _
           "Time Elapsed: " & Format(elapsedTime, "0.0") & " seconds", _
           vbInformation, "Repopulate Complete"
    
CleanExit:
    ' ========================================================================
    ' 8. RESTORE EXCEL SETTINGS
    ' ========================================================================
    ' IMPORTANT: Always restore these settings! If we don't, Excel will
    ' feel "broken" to the user (screen not updating, events not firing).
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    
    Set examinerLookup = Nothing
    Set repopSheet = Nothing
    Set thisWB = Nothing
    Set scheduleWB = Nothing
    
    Exit Sub
    
ErrorHandler:
    ' Log the error for debugging
    LogError "RepopulateModulesAndForms", Err.Number, Err.Description, _
             "Review: " & reviewNum
    
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error"
    
    ' Clean up any open workbook
    If Not scheduleWB Is Nothing Then
        scheduleWB.Close SaveChanges:=False
    End If
    
    Resume CleanExit
End Sub


' ============================================================================
' HELPER FUNCTIONS
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: BuildExaminerDictionary
' ----------------------------------------------------------------------------
' PURPOSE:
'   Creates a Dictionary mapping examiner numbers to examiner names.
'   This allows O(1) lookup instead of O(n) loop for each review.
'
' PARAMETERS:
'   examData - 2D array from the examiner lookup range (name in col 1, num in col 2)
'   dict     - Dictionary object to populate
'
' EXAMPLE:
'   After calling, dict("7") = "John Smith"
' ----------------------------------------------------------------------------
Private Sub BuildExaminerDictionary(ByRef examData As Variant, ByRef dict As Object)
    Dim i As Long
    Dim examNum As String
    Dim examName As String
    
    For i = 2 To UBound(examData, 1)  ' Start at 2 to skip header
        examNum = Trim(CStr(examData(i, 2)))
        examName = Trim(CStr(examData(i, 1)))
        
        If examNum <> "" And examName <> "" Then
            If Not dict.Exists(examNum) Then
                dict.Add examNum, examName
            End If
        End If
    Next i
End Sub

' ----------------------------------------------------------------------------
' Function: FindReviewWorkbook
' ----------------------------------------------------------------------------
' PURPOSE:
'   Searches for a review workbook file in the given directory and subdirectories.
'   Uses native Dir() function instead of clFileSearchModule class.
'
' PARAMETERS:
'   searchPath - Base directory to search (e.g., ...\TANF\)
'   reviewNum  - Review number to find
'   monthStr   - Sample month in YYYYMM format
'
' RETURNS:
'   String - Full path to the found file, or "" if not found
'
' SEARCH PATTERN:
'   Looks for: "Review Number {reviewNum} Month {monthStr} Examiner*.xls*"
'
' NOTE:
'   This is simpler than the old clFileSearchModule approach and has no
'   external dependencies. It searches subdirectories by walking the
'   folder structure.
' ----------------------------------------------------------------------------
Private Function FindReviewWorkbook(ByVal searchPath As String, _
                                     ByVal reviewNum As String, _
                                     ByVal monthStr As String) As String
    On Error Resume Next
    
    Dim fso As Object
    Dim folder As Object
    Dim subFolder As Object
    Dim file As Object
    Dim searchPattern As String
    Dim foundFile As String
    
    FindReviewWorkbook = ""
    
    ' Build the filename pattern we're looking for
    searchPattern = "Review Number " & reviewNum & " Month " & monthStr & " Examiner"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists(searchPath) Then
        Exit Function
    End If
    
    Set folder = fso.GetFolder(searchPath)
    
    ' Search subdirectories (Review Month folders)
    For Each subFolder In folder.SubFolders
        ' Check for case folders inside Review Month folders
        Dim caseFolder As Object
        For Each caseFolder In subFolder.SubFolders
            ' Look for the review file
            For Each file In caseFolder.Files
                If LCase(Right(file.Name, 4)) = ".xls" Or _
                   LCase(Right(file.Name, 5)) = ".xlsx" Or _
                   LCase(Right(file.Name, 5)) = ".xlsm" Then
                    If InStr(1, file.Name, searchPattern, vbTextCompare) > 0 Then
                        FindReviewWorkbook = file.path
                        Exit Function
                    End If
                End If
            Next file
        Next caseFolder
    Next subFolder
    
    Set fso = Nothing
End Function

' ----------------------------------------------------------------------------
' Function: GetProgramFolderNameForRepop
' ----------------------------------------------------------------------------
' PURPOSE:
'   Returns the folder name for a program type. This may differ from
'   GetProgramFolderName in Common_Utils if the folder naming convention
'   is different (e.g., "FS Positive" vs "SNAP Positive").
'
' PARAMETERS:
'   programType - The ProgramType enum value
'
' RETURNS:
'   String - The folder name used on the network
'
' NOTE:
'   The original code used "FS Positive" and "FS Negative" but some
'   installations may use "SNAP Positive" and "SNAP Negative".
'   Adjust these values to match your folder structure.
' ----------------------------------------------------------------------------
Private Function GetProgramFolderNameForRepop(ByVal programType As ProgramType) As String
    Select Case programType
        Case PROG_SNAP_POS
            GetProgramFolderNameForRepop = "FS Positive"  ' Or "SNAP Positive"
        Case PROG_SNAP_NEG
            GetProgramFolderNameForRepop = "FS Negative"  ' Or "SNAP Negative"
        Case PROG_TANF
            GetProgramFolderNameForRepop = "TANF"
        Case PROG_GA
            GetProgramFolderNameForRepop = "GA"
        Case PROG_MA_POS
            GetProgramFolderNameForRepop = "MA Positive"
        Case PROG_MA_NEG
            GetProgramFolderNameForRepop = "MA Negative"
        Case PROG_MA_PE
            GetProgramFolderNameForRepop = "MA PE"
        Case Else
            GetProgramFolderNameForRepop = ""
    End Select
End Function

' ----------------------------------------------------------------------------
' Sub: CopyModulesForProgram
' ----------------------------------------------------------------------------
' PURPOSE:
'   Copies the appropriate VBA modules from the source workbook to the
'   target workbook based on program type.
'
' PARAMETERS:
'   sourceWB - The workbook containing the modules to copy (this workbook)
'   targetWB - The workbook to copy modules into (the review schedule)
'   programType - The program type enum value
'
' MODULES BY PROGRAM:
'   See the Select Case below for which modules each program gets.
'   This replaces the long If/ElseIf chain in the original code.
' ----------------------------------------------------------------------------
Private Sub CopyModulesForProgram(ByRef sourceWB As Workbook, _
                                   ByRef targetWB As Workbook, _
                                   ByVal programType As ProgramType)
    ' Common modules that ALL programs get
    CopyModule sourceWB, "Common_Utils", targetWB
    CopyModule sourceWB, "Config_Settings", targetWB
    CopyModule sourceWB, "Review_Approval", targetWB
    CopyModule sourceWB, "Review_EditCheck", targetWB
    CopyModule sourceWB, "Review_FindingMemo", targetWB
    CopyModule sourceWB, "Review_CashMemos", targetWB
    CopyModule sourceWB, "Review_Appointments", targetWB
    
    ' Program-specific modules
    Select Case programType
        Case PROG_TANF
            CopyModule sourceWB, "Review_Drop", targetWB
            CopyModule sourceWB, "Review_TANF_Utils", targetWB
            
        Case PROG_GA
            CopyModule sourceWB, "Review_GA_Elements", targetWB
            
        Case PROG_SNAP_POS
            CopyModule sourceWB, "Review_Drop", targetWB
            CopyModule sourceWB, "Review_SNAP_Utils", targetWB
            CopyModule sourceWB, "Review_TANF_Utils", targetWB
            
        Case PROG_SNAP_NEG
            CopyModule sourceWB, "Review_TANF_Utils", targetWB
            
        Case PROG_MA_POS
            CopyModule sourceWB, "Review_Drop", targetWB
            CopyModule sourceWB, "Review_MA_Comp", targetWB
            CopyModule sourceWB, "Review_TANF_Utils", targetWB
            
        Case PROG_MA_NEG, PROG_MA_PE
            CopyModule sourceWB, "Review_TANF_Utils", targetWB
    End Select
End Sub

' ----------------------------------------------------------------------------
' Sub: CopyFormsForProgram
' ----------------------------------------------------------------------------
' PURPOSE:
'   Copies the appropriate UserForms from the source workbook to the
'   target workbook based on program type.
'
' PARAMETERS:
'   sourceWB - The workbook containing the forms to copy
'   targetWB - The workbook to copy forms into
'   programType - The program type enum value
' ----------------------------------------------------------------------------
Private Sub CopyFormsForProgram(ByRef sourceWB As Workbook, _
                                 ByRef targetWB As Workbook, _
                                 ByVal programType As ProgramType)
    ' Common forms that ALL programs get
    CopyForm sourceWB, "UF_SelectForms", targetWB
    CopyForm sourceWB, "UF_DatePicker", targetWB
    CopyForm sourceWB, "UF_TimePicker", targetWB
    
    ' Program-specific forms
    Select Case programType
        Case PROG_TANF
            CopyForm sourceWB, "UF_TANF_Results", targetWB
            CopyForm sourceWB, "UF_TANF_Helper", targetWB
            
        Case PROG_GA
            CopyForm sourceWB, "UF_GA_Helper1", targetWB
            CopyForm sourceWB, "UF_GA_Helper2", targetWB
            
        Case PROG_SNAP_POS
            ' Only common forms
            
        Case PROG_SNAP_NEG
            ' Only common forms
            
        Case PROG_MA_POS
            CopyForm sourceWB, "UF_MA_SelectForms", targetWB
            CopyForm sourceWB, "UF_MA_Comp2", targetWB
            CopyForm sourceWB, "UF_MA_Comp3", targetWB
            
        Case PROG_MA_NEG, PROG_MA_PE
            CopyForm sourceWB, "UF_MA_SelectForms", targetWB
    End Select
End Sub

' ----------------------------------------------------------------------------
' Sub: CopyModule
' ----------------------------------------------------------------------------
' PURPOSE:
'   Copies a VBA module from one workbook to another. If the module already
'   exists in the target, it's deleted first to ensure a clean replacement.
'
' PARAMETERS:
'   sourceWB   - Workbook containing the module
'   moduleName - Name of the module to copy
'   targetWB   - Workbook to copy the module into
'
' NOTE:
'   This requires "Trust access to the VBA project object model" to be
'   enabled in Excel's Trust Center settings.
' ----------------------------------------------------------------------------
Private Sub CopyModule(ByRef sourceWB As Workbook, _
                        ByVal moduleName As String, _
                        ByRef targetWB As Workbook)
    On Error Resume Next
    
    Dim sourceModule As Object
    Dim tempFile As String
    
    ' Check if module exists in source
    Set sourceModule = sourceWB.VBProject.VBComponents(moduleName)
    If sourceModule Is Nothing Then
        Exit Sub
    End If
    
    ' Delete existing module in target (if it exists)
    On Error Resume Next
    targetWB.VBProject.VBComponents.Remove _
        targetWB.VBProject.VBComponents(moduleName)
    Err.Clear
    On Error GoTo 0
    
    ' Export from source and import to target
    tempFile = Environ("TEMP") & "\" & moduleName & ".bas"
    
    On Error Resume Next
    sourceModule.Export tempFile
    If Err.Number = 0 Then
        targetWB.VBProject.VBComponents.Import tempFile
        Kill tempFile  ' Clean up temp file
    End If
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' Sub: CopyForm
' ----------------------------------------------------------------------------
' PURPOSE:
'   Copies a UserForm from one workbook to another. Similar to CopyModule
'   but for forms (.frm files).
'
' PARAMETERS:
'   sourceWB - Workbook containing the form
'   formName - Name of the form to copy
'   targetWB - Workbook to copy the form into
' ----------------------------------------------------------------------------
Private Sub CopyForm(ByRef sourceWB As Workbook, _
                      ByVal formName As String, _
                      ByRef targetWB As Workbook)
    On Error Resume Next
    
    Dim sourceForm As Object
    Dim tempFile As String
    
    ' Check if form exists in source
    Set sourceForm = sourceWB.VBProject.VBComponents(formName)
    If sourceForm Is Nothing Then
        Exit Sub
    End If
    
    ' Delete existing form in target (if it exists)
    On Error Resume Next
    targetWB.VBProject.VBComponents.Remove _
        targetWB.VBProject.VBComponents(formName)
    Err.Clear
    On Error GoTo 0
    
    ' Export from source and import to target
    tempFile = Environ("TEMP") & "\" & formName & ".frm"
    
    On Error Resume Next
    sourceForm.Export tempFile
    If Err.Number = 0 Then
        targetWB.VBProject.VBComponents.Import tempFile
        Kill tempFile  ' Clean up temp file
        ' Also delete .frx binary file if it exists
        If Dir(Environ("TEMP") & "\" & formName & ".frx") <> "" Then
            Kill Environ("TEMP") & "\" & formName & ".frx"
        End If
    End If
    On Error GoTo 0
End Sub


