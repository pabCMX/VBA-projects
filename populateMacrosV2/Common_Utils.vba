Attribute VB_Name = "Common_Utils"
' ============================================================================
' Common_Utils - Shared Utility Functions for QC Case Review System
' ============================================================================
' WHAT THIS MODULE DOES:
'   This module contains utility functions that are used throughout the entire
'   QC Case Review system. Instead of copying the same code into multiple
'   modules (which creates maintenance headaches), we put it here once and
'   call it from wherever it's needed.
'
' WHY THIS EXISTS:
'   The original codebase had the same code copied in 6-8 different places.
'   For example, the network drive detection code appeared in Module1, Module3,
'   repopulate_mod, and several others. When the network path changed, someone
'   had to find and update ALL copies - and they often missed some.
'
'   By centralizing these functions here:
'   - Changes only need to be made in ONE place
'   - Less chance of bugs from inconsistent updates
'   - Easier for new maintainers to understand the system
'
' FUNCTIONS IN THIS MODULE:
'   Network & Path Functions:
'     - GetDQCDriveLetter()       : Finds the mapped network drive
'     - GetDQCDriveLetterOrError(): Same as above, but shows error if not found
'     - PathExists()              : Checks if a file/folder exists
'     - ValidateNetworkPath()     : Validates path with error message
'
'   Program Detection Functions:
'     - GetProgramFromReviewNumber(): Determines program type from review #
'     - GetProgramFromSheetName()   : Determines program type from sheet name
'     - GetProgramFolderName()      : Gets folder name for a program type
'
'   Status & Routing Functions:
'     - DetermineStatusFolder()   : Figures out Clean/Error/Drop status
'     - StatusFolderToString()    : Converts status enum to folder name
'
'   Formatting Functions:
'     - FormatExaminerNumber()    : Formats examiner number without leading zero
'     - MonthNumberToName()       : Converts "01" to "January"
'     - LineNumberFormat()        : Formats line number as "01", "02", etc.
'     - RemoveLeadingZeros()      : Strips leading zeros from strings
'
'   Data Conversion Functions:
'     - ParseBISDate()            : Converts YYYYMMDD string to VBA Date
'     - SafeValue()               : Safely gets cell value with default
'
'   Logging & Error Functions:
'     - LogError()                : Writes error info to log file
'     - GetUserTimestamp()        : Gets "USERNAME MM/DD/YYYY" string
'
' HOW TO USE THIS MODULE:
'   From any other module, just call these functions directly:
'
'       Dim driveLetter As String
'       driveLetter = GetDQCDriveLetterOrError()
'       ' Now driveLetter contains something like "E:\DQC\"
'
'   Or:
'
'       Dim progType As ProgramType
'       progType = GetProgramFromReviewNumber("50012345")
'       ' progType = PROG_SNAP_POS (it's a SNAP Positive review)
'
' MAINTENANCE NOTES:
'   - If the network path changes, update GetDQCDriveLetter() only
'   - If new programs are added, update the ProgramType enum and related functions
'   - Keep this module focused on UTILITY functions only - no business logic!
'
' CHANGE LOG:
'   2026-01-02  Initial creation - extracted from 8+ duplicated implementations
' ============================================================================

Option Explicit

' ============================================================================
' PROGRAM TYPE ENUMERATION
' ============================================================================
' PURPOSE:
'   Provides a clear, readable way to identify program types throughout the code.
'   Instead of comparing strings like "SNAP Positive" everywhere (error-prone),
'   we use this enum. If you type PROG_SNAAP_POS by mistake, VBA catches it at
'   compile time instead of silently failing at runtime.
'
' VALUES:
'   PROG_UNKNOWN  = 0   : Review number doesn't match any known program
'   PROG_SNAP_POS = 1   : SNAP Positive reviews (prefixes: 50, 51, 55)
'   PROG_SNAP_NEG = 2   : SNAP Negative reviews (prefixes: 60, 61, 65, 66)
'   PROG_TANF     = 3   : TANF reviews (prefix: 14)
'   PROG_GA       = 4   : General Assistance reviews (prefix: 9)
'   PROG_MA_POS   = 5   : Medical Assistance Positive reviews (prefixes: 20, 21)
'   PROG_MA_NEG   = 6   : Medical Assistance Negative reviews (prefixes: 80-83)
'   PROG_MA_PE    = 7   : Medical Assistance PE reviews (prefix: 24)
'
' EXAMPLE:
'   Dim progType As ProgramType
'   progType = GetProgramFromReviewNumber("50012345")
'   If progType = PROG_SNAP_POS Then
'       ' Handle SNAP Positive case
'   End If
' ============================================================================
Public Enum ProgramType
    PROG_UNKNOWN = 0
    PROG_SNAP_POS = 1
    PROG_SNAP_NEG = 2
    PROG_TANF = 3
    PROG_GA = 4
    PROG_MA_POS = 5
    PROG_MA_NEG = 6
    PROG_MA_PE = 7
End Enum

' ============================================================================
' STATUS FOLDER ENUMERATION
' ============================================================================
' PURPOSE:
'   When a review is completed, it gets filed into a folder based on its status.
'   This enum represents those three possible outcomes.
'
' VALUES:
'   STATUS_UNKNOWN = 0  : Could not determine status (error condition)
'   STATUS_CLEAN   = 1  : Review found no errors - case was correct
'   STATUS_ERROR   = 2  : Review found errors - case had problems
'   STATUS_DROP    = 3  : Review was dropped (not completed for some reason)
'
' USED BY:
'   - DetermineStatusFolder() function below
'   - Approval workflow in Review_Approval module
' ============================================================================
Public Enum StatusFolder
    STATUS_UNKNOWN = 0
    STATUS_CLEAN = 1
    STATUS_ERROR = 2
    STATUS_DROP = 3
End Enum


' ============================================================================
' SECTION 1: NETWORK & PATH FUNCTIONS
' ============================================================================

' ----------------------------------------------------------------------------
' Function: GetDQCDriveLetter
' ----------------------------------------------------------------------------
' PURPOSE:
'   Finds which drive letter is mapped to the DQC network share. The DQC
'   (Data Quality Control) folder is where all schedules and review files live.
'
' HOW IT WORKS:
'   1. Gets a list of all mapped network drives on this computer
'   2. Loops through looking for the DQC server path
'   3. Returns the drive letter (like "E:") plus the DQC path
'
' RETURNS:
'   String - The full path to the DQC folder, e.g., "E:\DQC\"
'            Returns empty string "" if the drive isn't mapped
'
' IMPORTANT:
'   The network share can be mapped two ways:
'   - Mapped to "stat" folder: \\hsedcprapfpp001\oim\pwimdaubts04\data\stat
'     In this case we need to add "\DQC\" to the path
'   - Mapped directly to "dqc": \\hsedcprapfpp001\oim\pwimdaubts04\data\stat\dqc
'     In this case we just add "\" to the path
'
' EXAMPLE:
'   Dim drivePath As String
'   drivePath = GetDQCDriveLetter()
'   If drivePath = "" Then
'       MsgBox "Network drive not found!"
'   Else
'       ' drivePath = "E:\DQC\" or similar
'   End If
'
' WHY WE DO IT THIS WAY:
'   - Different users might have different drive letters mapped
'   - We can't hardcode "E:" because it might be "F:" on another computer
'   - The WScript.Network object gives us access to the drive mappings
' ----------------------------------------------------------------------------
Public Function GetDQCDriveLetter() As String
    On Error GoTo ErrorHandler
    
    ' These are the objects we use to access Windows networking info
    Dim WshNetwork As Object    ' Windows Script Host Network object
    Dim oDrives As Object       ' Collection of mapped drives
    Dim DUNC As String          ' The UNC path (\\server\share format)
    Dim DLetter As String       ' The drive letter we find
    Dim i As Long               ' Loop counter (use Long, not Integer!)
    
    ' Create the network object - this lets us see mapped drives
    Set WshNetwork = CreateObject("WScript.Network")
    Set oDrives = WshNetwork.EnumNetworkDrives
    
    ' oDrives is a collection where:
    '   - Even indices (0, 2, 4...) = drive letters (like "E:")
    '   - Odd indices (1, 3, 5...)  = UNC paths (like "\\server\share")
    ' So we step by 2 and check each pair
    
    DLetter = ""
    For i = 0 To oDrives.Count - 1 Step 2
        ' Get the UNC path for this drive (the +1 gives us the odd index)
        ' We use LCase so our comparison is case-insensitive
        DUNC = LCase("" & oDrives.Item(i + 1) & "")
        
        ' Check if this drive is mapped to either version of the DQC path
        If DUNC = "\\hsedcprapfpp001\oim\pwimdaubts04\data\stat" Then
            ' Mapped to "stat" folder - need to add \DQC\ to path
            DLetter = "" & oDrives.Item(i) & "\DQC\"
            Exit For
        ElseIf DUNC = "\\hsedcprapfpp001\oim\pwimdaubts04\data\stat\dqc" Then
            ' Mapped directly to "dqc" folder - just add \
            DLetter = "" & oDrives.Item(i) & "\"
            Exit For
        End If
    Next i
    
    GetDQCDriveLetter = DLetter
    
    ' Clean up objects to free memory
    Set oDrives = Nothing
    Set WshNetwork = Nothing
    Exit Function
    
ErrorHandler:
    ' If anything goes wrong, return empty string
    ' The caller can check for "" and handle appropriately
    GetDQCDriveLetter = ""
    Set oDrives = Nothing
    Set WshNetwork = Nothing
End Function

' ----------------------------------------------------------------------------
' Function: GetDQCDriveLetterOrError
' ----------------------------------------------------------------------------
' PURPOSE:
'   Same as GetDQCDriveLetter(), but automatically shows an error message
'   and stops execution if the drive isn't found. Use this version when
'   there's no point continuing without the network drive.
'
' RETURNS:
'   String - The full path to the DQC folder
'   NOTE: If drive not found, this function ENDS execution after showing error
'
' WHEN TO USE:
'   - Use GetDQCDriveLetter() if you want to handle the error yourself
'   - Use GetDQCDriveLetterOrError() if you want automatic error handling
'
' EXAMPLE:
'   ' This will either give you the path or stop the macro entirely
'   Dim drivePath As String
'   drivePath = GetDQCDriveLetterOrError()
'   ' If we get here, we definitely have a valid path
' ----------------------------------------------------------------------------
Public Function GetDQCDriveLetterOrError() As String
    Dim DLetter As String
    
    DLetter = GetDQCDriveLetter()
    
    If DLetter = "" Then
        ' Show error message with contact info
        MsgBox "Network Drive to DQC Directory is NOT correct" & Chr(13) & _
            "Contact Valerie or Nicole", vbCritical, "Network Error"
        ' End statement stops ALL code execution immediately
        End
    End If
    
    GetDQCDriveLetterOrError = DLetter
End Function

' ----------------------------------------------------------------------------
' Function: PathExists
' ----------------------------------------------------------------------------
' PURPOSE:
'   Checks whether a file or folder exists at the given path.
'   Simple but useful - avoids cryptic errors when trying to open
'   files that don't exist.
'
' PARAMETERS:
'   path - The file or folder path to check
'
' RETURNS:
'   Boolean - True if the path exists, False if not
'
' EXAMPLE:
'   If PathExists("E:\DQC\Schedules\") Then
'       ' Safe to use this folder
'   Else
'       MsgBox "Folder not found!"
'   End If
' ----------------------------------------------------------------------------
Public Function PathExists(ByVal path As String) As Boolean
    On Error Resume Next
    ' Dir() returns the filename if found, empty string if not
    PathExists = (Len(Dir(path, vbDirectory)) > 0)
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' Function: ValidateNetworkPath
' ----------------------------------------------------------------------------
' PURPOSE:
'   Checks if a path exists, and if not, shows a helpful error message.
'   Combines PathExists() with error messaging for convenience.
'
' PARAMETERS:
'   path     - The path to validate
'   pathDesc - Description of what this path is (for the error message)
'              Optional, defaults to "Path"
'
' RETURNS:
'   Boolean - True if path exists, False if not (error shown in this case)
'
' EXAMPLE:
'   If Not ValidateNetworkPath(examinerFolder, "Examiner Folder") Then
'       Exit Sub  ' Error already shown to user
'   End If
' ----------------------------------------------------------------------------
Public Function ValidateNetworkPath(ByVal path As String, _
                                    Optional ByVal pathDesc As String = "Path") As Boolean
    If Not PathExists(path) Then
        MsgBox pathDesc & ": " & path & " does NOT exist!!" & Chr(13) & _
            "Contact Valerie or Nicole", vbCritical, "Path Error"
        ValidateNetworkPath = False
    Else
        ValidateNetworkPath = True
    End If
End Function


' ============================================================================
' SECTION 2: PROGRAM DETECTION FUNCTIONS
' ============================================================================

' ----------------------------------------------------------------------------
' Function: GetProgramFromReviewNumber
' ----------------------------------------------------------------------------
' PURPOSE:
'   Determines which program a review belongs to based on its review number.
'   Review numbers have a prefix that indicates the program:
'     50, 51, 55     = SNAP Positive
'     60, 61, 65, 66 = SNAP Negative
'     14             = TANF
'     9              = GA (General Assistance)
'     20, 21         = MA Positive
'     24             = MA PE (Presumptive Eligibility)
'     80, 81, 82, 83 = MA Negative
'
' PARAMETERS:
'   reviewNumber - The review number as a string or number
'                  Can include leading zeros (they'll be stripped)
'
' RETURNS:
'   ProgramType enum value (PROG_SNAP_POS, PROG_TANF, etc.)
'   Returns PROG_UNKNOWN if the prefix doesn't match any program
'
' EXAMPLES:
'   GetProgramFromReviewNumber("50012345") returns PROG_SNAP_POS
'   GetProgramFromReviewNumber("14067890") returns PROG_TANF
'   GetProgramFromReviewNumber("9001234")  returns PROG_GA
'   GetProgramFromReviewNumber("99999999") returns PROG_UNKNOWN
'
' WHY WE CHECK 2 DIGITS FIRST:
'   Most programs use 2-digit prefixes (50, 14, 80, etc.), but GA uses
'   just "9". So we check 2-digit prefixes first, then fall back to
'   1-digit for GA. This avoids false matches.
' ----------------------------------------------------------------------------
Public Function GetProgramFromReviewNumber(ByVal reviewNumber As Variant) As ProgramType
    Dim reviewStr As String
    Dim prefix2 As String     ' First 2 characters
    Dim prefix1 As String     ' First 1 character
    
    reviewStr = CStr(reviewNumber)
    
    ' Remove leading zeros (e.g., "050012345" becomes "50012345")
    ' Some systems add leading zeros for formatting
    Do While Left(reviewStr, 1) = "0" And Len(reviewStr) > 1
        reviewStr = Mid(reviewStr, 2)
    Loop
    
    prefix2 = Left(reviewStr, 2)
    prefix1 = Left(reviewStr, 1)
    
    ' Check 2-digit prefixes first (most programs)
    Select Case prefix2
        Case "50", "51", "55"
            GetProgramFromReviewNumber = PROG_SNAP_POS
        Case "60", "61", "65", "66"
            GetProgramFromReviewNumber = PROG_SNAP_NEG
        Case "14"
            GetProgramFromReviewNumber = PROG_TANF
        Case "20", "21"
            GetProgramFromReviewNumber = PROG_MA_POS
        Case "24"
            GetProgramFromReviewNumber = PROG_MA_PE
        Case "80", "81", "82", "83"
            GetProgramFromReviewNumber = PROG_MA_NEG
        Case Else
            ' Check 1-digit prefix for GA
            If prefix1 = "9" Then
                GetProgramFromReviewNumber = PROG_GA
            Else
                GetProgramFromReviewNumber = PROG_UNKNOWN
            End If
    End Select
End Function

' ----------------------------------------------------------------------------
' Function: GetProgramFromSheetName
' ----------------------------------------------------------------------------
' PURPOSE:
'   Determines program type from a worksheet name. When a schedule is open,
'   the worksheet tab shows the review number. This function figures out
'   the program from that.
'
' PARAMETERS:
'   sheetName - The worksheet name (usually equals the review number)
'
' RETURNS:
'   ProgramType enum value
'
' NOTE:
'   This is essentially the same logic as GetProgramFromReviewNumber,
'   but kept separate in case sheet naming conventions differ from
'   review number formats in the future.
' ----------------------------------------------------------------------------
Public Function GetProgramFromSheetName(ByVal sheetName As String) As ProgramType
    Dim prefix1 As String
    Dim prefix2 As String
    
    prefix1 = Left(sheetName, 1)
    prefix2 = Left(sheetName, 2)
    
    Select Case prefix1
        Case "5"
            GetProgramFromSheetName = PROG_SNAP_POS
        Case "6"
            GetProgramFromSheetName = PROG_SNAP_NEG
        Case "1"
            GetProgramFromSheetName = PROG_TANF
        Case "9"
            GetProgramFromSheetName = PROG_GA
        Case "2"
            ' MA Positive vs MA PE - need to check 2nd digit
            If prefix2 = "24" Then
                GetProgramFromSheetName = PROG_MA_PE
            Else
                GetProgramFromSheetName = PROG_MA_POS
            End If
        Case "8"
            GetProgramFromSheetName = PROG_MA_NEG
        Case Else
            GetProgramFromSheetName = PROG_UNKNOWN
    End Select
End Function

' ----------------------------------------------------------------------------
' Function: GetProgramFolderName
' ----------------------------------------------------------------------------
' PURPOSE:
'   Converts a ProgramType enum value to the folder name used on the network.
'   The folder structure uses specific names for each program.
'
' PARAMETERS:
'   programType - The ProgramType enum value
'
' RETURNS:
'   String - The folder name (e.g., "SNAP Positive", "TANF")
'            Returns empty string for PROG_UNKNOWN
'
' FOLDER STRUCTURE:
'   E:\DQC\Schedules by Examiner Number\{Examiner Name} - {#}\{Program}\
'   Where {Program} is one of:
'     - "SNAP Positive"
'     - "SNAP Negative"
'     - "TANF"
'     - "GA"
'     - "MA Positive"
'     - "MA Negative"
'     - "MA PE"
' ----------------------------------------------------------------------------
Public Function GetProgramFolderName(ByVal programType As ProgramType) As String
    Select Case programType
        Case PROG_SNAP_POS
            GetProgramFolderName = "SNAP Positive"
        Case PROG_SNAP_NEG
            GetProgramFolderName = "SNAP Negative"
        Case PROG_TANF
            GetProgramFolderName = "TANF"
        Case PROG_GA
            GetProgramFolderName = "GA"
        Case PROG_MA_POS
            GetProgramFolderName = "MA Positive"
        Case PROG_MA_NEG
            GetProgramFolderName = "MA Negative"
        Case PROG_MA_PE
            GetProgramFolderName = "MA PE"
        Case Else
            GetProgramFolderName = ""
    End Select
End Function


' ============================================================================
' SECTION 3: STATUS & ROUTING FUNCTIONS
' ============================================================================

' ----------------------------------------------------------------------------
' Function: DetermineStatusFolder
' ----------------------------------------------------------------------------
' PURPOSE:
'   Determines whether a completed review should be filed as Clean, Error,
'   or Drop. Each program has different rules for making this determination
'   based on specific cells in the schedule worksheet.
'
' PARAMETERS:
'   programType - The ProgramType enum value
'   ws          - The worksheet to read status values from
'
' RETURNS:
'   StatusFolder enum value (STATUS_CLEAN, STATUS_ERROR, STATUS_DROP)
'   Returns STATUS_UNKNOWN if something goes wrong
'
' PROGRAM-SPECIFIC LOGIC:
'
'   TANF & GA:
'     - AI10 > 1 = Drop (disposition code indicates case was dropped)
'     - AL10 = 1 = Clean (review findings = no error)
'     - Otherwise = Error
'
'   SNAP Positive:
'     - C22 > 1 = Drop
'     - K22 = 1 = Clean
'     - Otherwise = Error
'
'   SNAP Negative:
'     - F29 > 1 = Drop
'     - E47 has value = Error (error element is populated)
'     - Otherwise = Clean
'
'   MA Positive:
'     - F16 <> 1 = Drop (initial eligibility check failed)
'     - F16 = 1 AND S16 = 1 = Clean
'     - Otherwise = Error
'
'   MA Negative / MA PE:
'     - Complex logic involving M56, C25, C40, C56 cells
'     - See inline comments below for details
'
' WHY EACH PROGRAM IS DIFFERENT:
'   Each program's schedule has a different layout with status indicators
'   in different cells. These cell references match the schedule templates.
'   If the templates change, update the cell references here.
' ----------------------------------------------------------------------------
Public Function DetermineStatusFolder(ByVal programType As ProgramType, _
                                       ByVal ws As Worksheet) As StatusFolder
    On Error GoTo ErrorHandler
    
    Select Case programType
    
        ' ----------------------------------------------------------------
        ' TANF and GA use the same logic
        ' ----------------------------------------------------------------
        Case PROG_TANF, PROG_GA
            ' AI10 = Disposition Code (1 = complete, >1 = dropped)
            ' AL10 = Review Findings (1 = clean, >1 = error)
            If ws.Range("AI10") > 1 Then
                DetermineStatusFolder = STATUS_DROP
            ElseIf ws.Range("AL10") = 1 Then
                DetermineStatusFolder = STATUS_CLEAN
            Else
                DetermineStatusFolder = STATUS_ERROR
            End If
            
        ' ----------------------------------------------------------------
        ' SNAP Positive
        ' ----------------------------------------------------------------
        Case PROG_SNAP_POS
            ' C22 = Disposition Code
            ' K22 = Finding Code
            If ws.Range("C22") > 1 Then
                DetermineStatusFolder = STATUS_DROP
            ElseIf ws.Range("K22") = 1 Then
                DetermineStatusFolder = STATUS_CLEAN
            Else
                DetermineStatusFolder = STATUS_ERROR
            End If
            
        ' ----------------------------------------------------------------
        ' SNAP Negative
        ' ----------------------------------------------------------------
        Case PROG_SNAP_NEG
            ' F29 = Disposition Code
            ' E47 = Error element (if populated, there's an error)
            If ws.Range("F29") > 1 Then
                DetermineStatusFolder = STATUS_DROP
            ElseIf ws.Range("E47") <> "" Then
                DetermineStatusFolder = STATUS_ERROR
            Else
                DetermineStatusFolder = STATUS_CLEAN
            End If
            
        ' ----------------------------------------------------------------
        ' MA Positive
        ' ----------------------------------------------------------------
        Case PROG_MA_POS
            ' F16 = Initial Eligibility (1 = eligible)
            ' S16 = Review status
            If ws.Range("F16") <> 1 Then
                DetermineStatusFolder = STATUS_DROP
            ElseIf ws.Range("F16") = 1 And Val(ws.Range("S16")) = 1 Then
                DetermineStatusFolder = STATUS_CLEAN
            Else
                DetermineStatusFolder = STATUS_ERROR
            End If
            
        ' ----------------------------------------------------------------
        ' MA Negative and MA PE
        ' ----------------------------------------------------------------
        Case PROG_MA_NEG, PROG_MA_PE
            ' Complex logic for MA Negative:
            ' M56 = Disposition code (non-zero = Drop)
            ' C25 = Hearing request code (2-4 = Error condition)
            ' C40 = Eligibility requirement code (2 = Error, 3-5 with C56=2 = Error)
            ' C56 = Field investigation code
            
            If ws.Range("M56") <> 0 Then
                DetermineStatusFolder = STATUS_DROP
            ElseIf (ws.Range("C25") > 1 And ws.Range("C25") < 5) Or _
                   ws.Range("C40") = 2 Or _
                   ((ws.Range("C40") > 2 And ws.Range("C40") < 6) And ws.Range("C56") = 2) Then
                DetermineStatusFolder = STATUS_ERROR
            Else
                DetermineStatusFolder = STATUS_CLEAN
            End If
            
        Case Else
            DetermineStatusFolder = STATUS_UNKNOWN
    End Select
    
    Exit Function
    
ErrorHandler:
    ' If we can't read the cells for some reason, return Unknown
    DetermineStatusFolder = STATUS_UNKNOWN
End Function

' ----------------------------------------------------------------------------
' Function: StatusFolderToString
' ----------------------------------------------------------------------------
' PURPOSE:
'   Converts a StatusFolder enum value to the actual folder name string.
'   Used when building file paths for saving completed reviews.
'
' PARAMETERS:
'   status - The StatusFolder enum value
'
' RETURNS:
'   String - "Clean", "Error", "Drop", or "" for unknown
'
' EXAMPLE:
'   Dim status As StatusFolder
'   status = DetermineStatusFolder(PROG_TANF, ActiveSheet)
'   savePath = basePath & StatusFolderToString(status) & "\"
'   ' savePath might be "E:\DQC\...\Clean\" or "E:\DQC\...\Error\"
' ----------------------------------------------------------------------------
Public Function StatusFolderToString(ByVal status As StatusFolder) As String
    Select Case status
        Case STATUS_CLEAN
            StatusFolderToString = "Clean"
        Case STATUS_ERROR
            StatusFolderToString = "Error"
        Case STATUS_DROP
            StatusFolderToString = "Drop"
        Case Else
            StatusFolderToString = ""
    End Select
End Function


' ============================================================================
' SECTION 4: FORMATTING FUNCTIONS
' ============================================================================

' ----------------------------------------------------------------------------
' Function: FormatExaminerNumber
' ----------------------------------------------------------------------------
' PURPOSE:
'   Formats an examiner number by removing the leading zero if present.
'   Examiner numbers are stored as 2-digit strings (01, 02, ... 12) but
'   folder names use single digits for 1-9 (like "Jane Smith - 1").
'
' PARAMETERS:
'   examNum - The examiner number (can be string or number)
'
' RETURNS:
'   String - Formatted examiner number without leading zero
'
' EXAMPLES:
'   FormatExaminerNumber("01") returns "1"
'   FormatExaminerNumber("07") returns "7"
'   FormatExaminerNumber("12") returns "12"
'   FormatExaminerNumber(5)    returns "5"
' ----------------------------------------------------------------------------
Public Function FormatExaminerNumber(ByVal examNum As Variant) As String
    Dim exnumstr As String
    
    ' Format as 2-digit string first (handles both string and number input)
    exnumstr = Application.WorksheetFunction.Text(examNum, "00")
    
    ' Remove leading zero if present
    If Left(exnumstr, 1) = "0" Then
        exnumstr = Right(exnumstr, 1)
    End If
    
    FormatExaminerNumber = exnumstr
End Function

' ----------------------------------------------------------------------------
' Function: MonthNumberToName
' ----------------------------------------------------------------------------
' PURPOSE:
'   Converts a month number (1-12 or "01"-"12") to the full month name.
'   Used when building folder paths like "Review Month January 2024".
'
' PARAMETERS:
'   monthNum - Month number as string ("01"-"12") or integer (1-12)
'
' RETURNS:
'   String - Full month name, or empty string if invalid
'
' EXAMPLES:
'   MonthNumberToName("01")  returns "January"
'   MonthNumberToName(7)     returns "July"
'   MonthNumberToName("12")  returns "December"
'   MonthNumberToName("13")  returns ""
' ----------------------------------------------------------------------------
Public Function MonthNumberToName(ByVal monthNum As Variant) As String
    Dim monthStr As String
    
    ' Handle both integer and string input - normalize to 2-digit string
    If IsNumeric(monthNum) Then
        monthStr = Format(CInt(monthNum), "00")
    Else
        monthStr = Right("00" & CStr(monthNum), 2)
    End If
    
    Select Case monthStr
        Case "01": MonthNumberToName = "January"
        Case "02": MonthNumberToName = "February"
        Case "03": MonthNumberToName = "March"
        Case "04": MonthNumberToName = "April"
        Case "05": MonthNumberToName = "May"
        Case "06": MonthNumberToName = "June"
        Case "07": MonthNumberToName = "July"
        Case "08": MonthNumberToName = "August"
        Case "09": MonthNumberToName = "September"
        Case "10": MonthNumberToName = "October"
        Case "11": MonthNumberToName = "November"
        Case "12": MonthNumberToName = "December"
        Case Else: MonthNumberToName = ""
    End Select
End Function

' ----------------------------------------------------------------------------
' Function: LineNumberFormat
' ----------------------------------------------------------------------------
' PURPOSE:
'   Formats a line number as a 2-digit string with leading zero.
'   Used throughout the population modules when writing data to schedules.
'
' PARAMETERS:
'   lineNum - Line number as variant (string or number)
'
' RETURNS:
'   String - Formatted line number ("01", "02", ... "12")
'
' EXAMPLES:
'   LineNumberFormat(1)   returns "01"
'   LineNumberFormat(9)   returns "09"
'   LineNumberFormat(12)  returns "12"
' ----------------------------------------------------------------------------
Public Function LineNumberFormat(ByVal lineNum As Variant) As String
    LineNumberFormat = Application.WorksheetFunction.Text(lineNum, "00")
End Function

' ----------------------------------------------------------------------------
' Function: RemoveLeadingZeros
' ----------------------------------------------------------------------------
' PURPOSE:
'   Strips all leading zeros from a string. Keeps at least one character
'   (so "000" becomes "0", not empty string).
'
' PARAMETERS:
'   value - String value to process
'
' RETURNS:
'   String - Value without leading zeros
'
' EXAMPLES:
'   RemoveLeadingZeros("00123")  returns "123"
'   RemoveLeadingZeros("007")    returns "7"
'   RemoveLeadingZeros("000")    returns "0"
'   RemoveLeadingZeros("100")    returns "100"
' ----------------------------------------------------------------------------
Public Function RemoveLeadingZeros(ByVal value As String) As String
    Do While Left(value, 1) = "0" And Len(value) > 1
        value = Mid(value, 2)
    Loop
    RemoveLeadingZeros = value
End Function


' ============================================================================
' SECTION 5: DATA CONVERSION FUNCTIONS
' ============================================================================

' ----------------------------------------------------------------------------
' Function: ParseBISDate
' ----------------------------------------------------------------------------
' PURPOSE:
'   Converts a date from BIS format (YYYYMMDD) to a VBA Date type.
'   BIS (Benefits Information System) exports dates as 8-digit numbers
'   like 20240115 for January 15, 2024.
'
' PARAMETERS:
'   dateString - Date in YYYYMMDD format (can be string or number)
'
' RETURNS:
'   Date - VBA Date value
'   Returns 0 if the input is empty, invalid, or can't be parsed
'
' EXAMPLES:
'   ParseBISDate("20240115")  returns #1/15/2024#
'   ParseBISDate(20231225)    returns #12/25/2023#
'   ParseBISDate("")          returns 0
'   ParseBISDate("invalid")   returns 0
'
' WHY THIS IS NEEDED:
'   VBA's built-in date functions don't recognize YYYYMMDD format.
'   CDate("20240115") would fail or give wrong results.
' ----------------------------------------------------------------------------
Public Function ParseBISDate(ByVal dateString As Variant) As Date
    On Error GoTo ErrorHandler
    
    Dim ds As String
    Dim yyyy As Long
    Dim mm As Long
    Dim dd As Long
    
    ds = CStr(dateString)
    
    ' Check for empty or too short
    If Len(ds) < 8 Or Val(ds) = 0 Then
        ParseBISDate = 0
        Exit Function
    End If
    
    ' Parse YYYYMMDD format
    yyyy = Val(Left(ds, 4))          ' Characters 1-4 = year
    mm = Val(Mid(ds, 5, 2))          ' Characters 5-6 = month
    dd = Val(Right(ds, 2))           ' Characters 7-8 = day
    
    ' Basic validation
    If yyyy < 1900 Or yyyy > 2100 Then GoTo ErrorHandler
    If mm < 1 Or mm > 12 Then GoTo ErrorHandler
    If dd < 1 Or dd > 31 Then GoTo ErrorHandler
    
    ' DateSerial builds a valid VBA date from year, month, day
    ParseBISDate = DateSerial(yyyy, mm, dd)
    Exit Function
    
ErrorHandler:
    ParseBISDate = 0
End Function

' ----------------------------------------------------------------------------
' Function: SafeValue
' ----------------------------------------------------------------------------
' PURPOSE:
'   Safely retrieves a cell's value, returning a default if the cell is
'   empty, contains an error, or can't be read for any reason.
'
' PARAMETERS:
'   rng        - The Range (cell) to get the value from
'   defaultVal - Value to return if cell is empty or has an error
'                Optional, defaults to empty string ""
'
' RETURNS:
'   Variant - The cell's value, or the default value
'
' EXAMPLE:
'   ' Instead of this (crashes on empty/error cells):
'   myValue = Range("A1").Value
'
'   ' Use this (safely handles problems):
'   myValue = SafeValue(Range("A1"), "N/A")
'
' WHY THIS IS USEFUL:
'   - Range.Value can throw errors if the cell contains an error like #N/A
'   - Empty cells might need to be treated as a specific default
'   - This avoids scattered On Error Resume Next throughout the code
' ----------------------------------------------------------------------------
Public Function SafeValue(ByVal rng As Range, _
                         Optional ByVal defaultVal As Variant = "") As Variant
    On Error Resume Next
    SafeValue = rng.Value
    If Err.Number <> 0 Or IsEmpty(SafeValue) Then
        SafeValue = defaultVal
    End If
    On Error GoTo 0
End Function


' ============================================================================
' SECTION 6: WORD DOCUMENT AND MAIL MERGE FUNCTIONS
' ============================================================================

' ----------------------------------------------------------------------------
' Function: OpenWordTemplate
' ----------------------------------------------------------------------------
' PURPOSE:
'   Opens a Word document template for mail merge operations. Creates a Word
'   application instance if one doesn't exist, or uses an existing one.
'
' PARAMETERS:
'   templatePath - Full path to the Word template file (.docx or .dotx)
'
' RETURNS:
'   Object - The Word Document object, or Nothing if failed
'
' EXAMPLE:
'   Dim wdDoc As Object
'   Set wdDoc = OpenWordTemplate("E:\DQC\Templates\FindingsMemo.docx")
'   If Not wdDoc Is Nothing Then
'       ' Work with the document
'   End If
'
' NOTE:
'   Caller is responsible for setting Word.Application.Visible = True
'   if they want the user to see the document.
' ----------------------------------------------------------------------------
Public Function OpenWordTemplate(ByVal templatePath As String) As Object
    On Error GoTo ErrorHandler
    
    Dim wdApp As Object
    Dim wdDoc As Object
    
    ' Check if file exists first
    If Not PathExists(templatePath) Then
        LogError "OpenWordTemplate", 0, "Template not found", templatePath
        Set OpenWordTemplate = Nothing
        Exit Function
    End If
    
    ' Try to get existing Word instance, or create new one
    On Error Resume Next
    Set wdApp = GetObject(, "Word.Application")
    If Err.Number <> 0 Then
        Err.Clear
        Set wdApp = CreateObject("Word.Application")
    End If
    On Error GoTo ErrorHandler
    
    ' Open the template
    Set wdDoc = wdApp.Documents.Open(templatePath)
    
    Set OpenWordTemplate = wdDoc
    Exit Function
    
ErrorHandler:
    LogError "OpenWordTemplate", Err.Number, Err.Description, templatePath
    Set OpenWordTemplate = Nothing
End Function

' ----------------------------------------------------------------------------
' Sub: ExecuteMailMerge
' ----------------------------------------------------------------------------
' PURPOSE:
'   Performs a Word mail merge by replacing placeholder text with values.
'   This is used to populate template documents with case data.
'
' PARAMETERS:
'   wdDoc    - The Word Document object
'   fieldMap - A Dictionary of placeholder->value pairs
'              Keys are the placeholder text (e.g., "<<CLIENT_NAME>>")
'              Values are the replacement text
'
' EXAMPLE:
'   Dim fields As Object
'   Set fields = CreateObject("Scripting.Dictionary")
'   fields.Add "<<CLIENT_NAME>>", "John Smith"
'   fields.Add "<<CASE_NUMBER>>", "12345678"
'   Call ExecuteMailMerge(wdDoc, fields)
' ----------------------------------------------------------------------------
Public Sub ExecuteMailMerge(ByRef wdDoc As Object, ByRef fieldMap As Object)
    On Error GoTo ErrorHandler
    
    Dim key As Variant
    Dim findObj As Object
    
    ' Iterate through all placeholders in the map
    For Each key In fieldMap.Keys
        With wdDoc.Content.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = CStr(key)
            .Replacement.Text = CStr(fieldMap(key))
            .Forward = True
            .Wrap = 1  ' wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .Execute Replace:=2  ' wdReplaceAll
        End With
    Next key
    
    Exit Sub
    
ErrorHandler:
    LogError "ExecuteMailMerge", Err.Number, Err.Description, ""
End Sub


' ============================================================================
' SECTION 7: EXAMINER AND PERSONNEL LOOKUP FUNCTIONS
' ============================================================================

' ----------------------------------------------------------------------------
' Function: GetExaminerInfo
' ----------------------------------------------------------------------------
' PURPOSE:
'   Retrieves examiner information from the lookup table based on examiner
'   number. Returns a Dictionary containing name, email, supervisor, etc.
'
' PARAMETERS:
'   examinerNum - The examiner number (1-12 typically)
'   dqcPath     - Optional: The DQC drive path (will be detected if omitted)
'
' RETURNS:
'   Object - A Scripting.Dictionary with keys:
'            "Name"       - Examiner's full name
'            "Email"      - Examiner's email address
'            "Supervisor" - Supervisor's name
'            "SupEmail"   - Supervisor's email
'            "Region"     - Region assignment
'            Returns Nothing if examiner not found
'
' EXAMPLE:
'   Dim examInfo As Object
'   Set examInfo = GetExaminerInfo("05")
'   If Not examInfo Is Nothing Then
'       MsgBox "Examiner: " & examInfo("Name")
'   End If
'
' NOTE:
'   Reads from the "Finding Memo Data Source.xlsx" file on the network.
' ----------------------------------------------------------------------------
Public Function GetExaminerInfo(ByVal examinerNum As Variant, _
                                Optional ByVal dqcPath As String = "") As Object
    On Error GoTo ErrorHandler
    
    Dim examDict As Object
    Dim dsWB As Workbook
    Dim dsWS As Worksheet
    Dim examNumFormatted As String
    Dim dsPath As String
    Dim i As Long
    
    Set examDict = CreateObject("Scripting.Dictionary")
    
    ' Get DQC path if not provided
    If dqcPath = "" Then
        dqcPath = GetDQCDriveLetter()
        If dqcPath = "" Then
            Set GetExaminerInfo = Nothing
            Exit Function
        End If
    End If
    
    ' Format examiner number
    examNumFormatted = FormatExaminerNumber(examinerNum)
    
    ' Build path to data source
    dsPath = dqcPath & "Finding Memo\Finding Memo Data Source.xlsx"
    
    If Not PathExists(dsPath) Then
        Set GetExaminerInfo = Nothing
        Exit Function
    End If
    
    ' Open data source (read-only, no update links)
    Application.ScreenUpdating = False
    Set dsWB = Workbooks.Open(Filename:=dsPath, UpdateLinks:=False, ReadOnly:=True)
    Set dsWS = dsWB.Worksheets(1)
    
    ' Look up examiner in column A, data in columns B-F
    For i = 2 To 20  ' Assuming max 20 examiners
        If CStr(dsWS.Cells(i, 1).Value) = examNumFormatted Then
            examDict.Add "Name", dsWS.Cells(i, 2).Value
            examDict.Add "Email", dsWS.Cells(i, 3).Value
            examDict.Add "Supervisor", dsWS.Cells(i, 4).Value
            examDict.Add "SupEmail", dsWS.Cells(i, 5).Value
            examDict.Add "Region", dsWS.Cells(i, 6).Value
            Exit For
        End If
    Next i
    
    ' Close data source without saving
    dsWB.Close False
    Application.ScreenUpdating = True
    
    If examDict.Count > 0 Then
        Set GetExaminerInfo = examDict
    Else
        Set GetExaminerInfo = Nothing
    End If
    
    Exit Function
    
ErrorHandler:
    On Error Resume Next
    If Not dsWB Is Nothing Then dsWB.Close False
    Application.ScreenUpdating = True
    Set GetExaminerInfo = Nothing
End Function


' ============================================================================
' SECTION 8: ADDRESS AND TEXT FORMATTING FUNCTIONS
' ============================================================================

' ----------------------------------------------------------------------------
' Function: FormatAddress
' ----------------------------------------------------------------------------
' PURPOSE:
'   Standardizes address formatting from BIS data. Handles the conversion
'   of uppercase addresses to proper case and combines address lines.
'
' PARAMETERS:
'   street1    - First street address line
'   street2    - Second street address line (apartment, etc.)
'   city       - City name
'   stateZip   - State and ZIP code (may be combined or separate)
'
' RETURNS:
'   String - Formatted address with proper capitalization
'
' EXAMPLE:
'   addr = FormatAddress("123 MAIN ST", "APT 4", "HARRISBURG", "PA 17101")
'   ' Returns "123 Main St, Apt 4" & vbCrLf & "Harrisburg, PA 17101"
' ----------------------------------------------------------------------------
Public Function FormatAddress(ByVal street1 As String, _
                              ByVal street2 As String, _
                              ByVal city As String, _
                              ByVal stateZip As String) As String
    Dim result As String
    Dim cityPart As String
    Dim statePart As String
    Dim tempArr() As String
    
    ' Format street line 1 (proper case)
    street1 = StrConv(Trim(street1), vbProperCase)
    
    ' Add street line 2 if present
    If Trim(street2) <> "" Then
        street2 = StrConv(Trim(street2), vbProperCase)
        result = street1 & ", " & street2
    Else
        result = street1
    End If
    
    ' Format city (proper case)
    cityPart = StrConv(Trim(city), vbProperCase)
    
    ' Parse state and ZIP if combined with comma
    If InStr(stateZip, ",") > 0 Then
        tempArr = Split(stateZip, ",")
        cityPart = StrConv(Trim(tempArr(LBound(tempArr))), vbProperCase)
        statePart = Trim(tempArr(UBound(tempArr)))
    Else
        statePart = Trim(stateZip)
    End If
    
    ' Build final address
    result = result & vbCrLf & cityPart & ", " & statePart
    
    FormatAddress = result
End Function

' ----------------------------------------------------------------------------
' Function: FormatClientName
' ----------------------------------------------------------------------------
' PURPOSE:
'   Formats a client name from uppercase BIS format to proper case.
'
' PARAMETERS:
'   rawName - The name in uppercase (e.g., "SMITH, JOHN A")
'
' RETURNS:
'   String - Properly formatted name (e.g., "Smith, John A")
' ----------------------------------------------------------------------------
Public Function FormatClientName(ByVal rawName As String) As String
    FormatClientName = StrConv(Trim(rawName), vbProperCase)
End Function


' ============================================================================
' SECTION 9: DATA EXTRACTION FUNCTIONS
' ============================================================================

' ----------------------------------------------------------------------------
' Function: ExtractReviewData
' ----------------------------------------------------------------------------
' PURPOSE:
'   Batch-reads schedule data into a 2D array for efficient access.
'   This is much faster than reading cells one at a time when you need
'   multiple values from the same range.
'
' PARAMETERS:
'   ws        - The worksheet to read from
'   rangeAddr - The range address to read (e.g., "A1:AQ100")
'
' RETURNS:
'   Variant - A 2D array containing the cell values
'             Access as: dataArray(row, col) where row/col are 1-based
'             Returns Empty if the read fails
'
' EXAMPLE:
'   Dim scheduleData As Variant
'   scheduleData = ExtractReviewData(ActiveSheet, "A1:AQ100")
'   If Not IsEmpty(scheduleData) Then
'       clientName = scheduleData(4, 2)  ' Row 4, Column B
'       reviewNum = scheduleData(18, 1)  ' Row 18, Column A
'   End If
'
' PERFORMANCE:
'   Reading a 100x50 range as an array: ~1ms
'   Reading same range cell-by-cell:    ~500ms
' ----------------------------------------------------------------------------
Public Function ExtractReviewData(ByRef ws As Worksheet, _
                                   ByVal rangeAddr As String) As Variant
    On Error GoTo ErrorHandler
    
    ExtractReviewData = ws.Range(rangeAddr).Value
    Exit Function
    
ErrorHandler:
    ExtractReviewData = Empty
End Function

' ----------------------------------------------------------------------------
' Function: GetCellFromArray
' ----------------------------------------------------------------------------
' PURPOSE:
'   Safely retrieves a value from a 2D data array with bounds checking.
'   Returns a default value if the indices are out of bounds.
'
' PARAMETERS:
'   dataArray  - The 2D array from ExtractReviewData
'   row        - Row index (1-based)
'   col        - Column index (1-based)
'   defaultVal - Value to return if out of bounds or empty
'
' RETURNS:
'   Variant - The cell value or default value
'
' EXAMPLE:
'   phone = GetCellFromArray(scheduleData, 5, 36, "")  ' Get AJ5, default ""
' ----------------------------------------------------------------------------
Public Function GetCellFromArray(ByRef dataArray As Variant, _
                                  ByVal row As Long, _
                                  ByVal col As Long, _
                                  Optional ByVal defaultVal As Variant = "") As Variant
    On Error GoTo ReturnDefault
    
    ' Check if array is valid
    If IsEmpty(dataArray) Then GoTo ReturnDefault
    
    ' Check bounds
    If row < LBound(dataArray, 1) Or row > UBound(dataArray, 1) Then GoTo ReturnDefault
    If col < LBound(dataArray, 2) Or col > UBound(dataArray, 2) Then GoTo ReturnDefault
    
    ' Get value
    If IsEmpty(dataArray(row, col)) Then
        GetCellFromArray = defaultVal
    Else
        GetCellFromArray = dataArray(row, col)
    End If
    Exit Function
    
ReturnDefault:
    GetCellFromArray = defaultVal
End Function

' ----------------------------------------------------------------------------
' Function: ColumnLetterToNumber
' ----------------------------------------------------------------------------
' PURPOSE:
'   Converts a column letter (A, B, ..., Z, AA, AB, ...) to its numeric index.
'   Useful when you have cell references like "AJ5" and need the column number.
'
' PARAMETERS:
'   colLetter - The column letter(s)
'
' RETURNS:
'   Long - The column number (1 for A, 2 for B, 27 for AA, etc.)
' ----------------------------------------------------------------------------
Public Function ColumnLetterToNumber(ByVal colLetter As String) As Long
    ColumnLetterToNumber = Range(colLetter & "1").Column
End Function

' ----------------------------------------------------------------------------
' Function: ColumnNumberToLetter
' ----------------------------------------------------------------------------
' PURPOSE:
'   Converts a column number to its letter representation.
'
' PARAMETERS:
'   colNum - The column number (1-based)
'
' RETURNS:
'   String - The column letter(s)
' ----------------------------------------------------------------------------
Public Function ColumnNumberToLetter(ByVal colNum As Long) As String
    ColumnNumberToLetter = Split(Cells(1, colNum).Address, "$")(1)
End Function


' ============================================================================
' SECTION 10: LOGGING & ERROR FUNCTIONS
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: LogError
' ----------------------------------------------------------------------------
' PURPOSE:
'   Writes error information to a log file for debugging. This is especially
'   useful for tracking down problems that happen on user machines where
'   we can't watch the code run.
'
' PARAMETERS:
'   procedureName - Name of the sub/function where the error occurred
'   errNum        - VBA error number (from Err.Number)
'   errDesc       - VBA error description (from Err.Description)
'   contextInfo   - Optional additional info (like which file was being processed)
'
' LOG FILE LOCATION:
'   Created in the same folder as the workbook containing this code.
'   Named "ErrorLog_YYYYMMDD.txt" (new file each day)
'
' LOG FILE FORMAT:
'   Date/Time | Procedure | ErrorNum | Description | Context | Username
'
' EXAMPLE:
'   Sub MySub()
'       On Error GoTo ErrorHandler
'       ' ... code that might fail ...
'       Exit Sub
'   ErrorHandler:
'       LogError "MySub", Err.Number, Err.Description, "Processing file: " & fileName
'       MsgBox "An error occurred - see log for details"
'   End Sub
'
' WHY WE LOG ERRORS:
'   - Users often can't describe exactly what happened
'   - The log captures technical details for debugging
'   - Daily files prevent the log from growing forever
'   - Includes username to identify who had the problem
' ----------------------------------------------------------------------------
Public Sub LogError(ByVal procedureName As String, _
                    ByVal errNum As Long, _
                    ByVal errDesc As String, _
                    Optional ByVal contextInfo As String = "")
    On Error Resume Next  ' Don't let logging errors cause more errors!
    
    Dim logFile As Integer
    Dim logPath As String
    Dim logEntry As String
    
    ' Create log path in same folder as workbook, with today's date
    logPath = ThisWorkbook.path & "\ErrorLog_" & Format(Date, "YYYYMMDD") & ".txt"
    
    ' Build log entry as pipe-delimited string
    ' (pipes are unlikely to appear in error messages, unlike commas)
    logEntry = Format(Now, "YYYY-MM-DD HH:NN:SS") & "|" & _
               procedureName & "|" & _
               errNum & "|" & _
               errDesc & "|" & _
               contextInfo & "|" & _
               Environ("USERNAME") & vbCrLf
    
    ' Append to log file (creates file if it doesn't exist)
    logFile = FreeFile
    Open logPath For Append As #logFile
    Print #logFile, logEntry
    Close #logFile
    
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' Function: GetUserTimestamp
' ----------------------------------------------------------------------------
' PURPOSE:
'   Creates a timestamp string with the current Windows username and date.
'   Used for approval fields to record who approved and when.
'
' RETURNS:
'   String - Format "USERNAME MM/DD/YYYY"
'            Example: "JSMITH 01/15/2024"
'
' EXAMPLE:
'   Range("A1").Value = GetUserTimestamp()
'   ' Cell now contains "JSMITH 01/15/2024"
'
' WHY ENVIRON("USERNAME"):
'   This gets the Windows login name directly from the environment.
'   It's more reliable than Application.UserName (which users can change).
' ----------------------------------------------------------------------------
Public Function GetUserTimestamp() As String
    GetUserTimestamp = Environ("USERNAME") & " " & Date
End Function


