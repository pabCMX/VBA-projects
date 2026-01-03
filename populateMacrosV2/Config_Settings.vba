Attribute VB_Name = "Config_Settings"
' ============================================================================
' Config_Settings - Centralized Configuration for QC Case Review System
' ============================================================================
' WHAT THIS MODULE DOES:
'   This module stores all the "magic numbers" and hardcoded values used
'   throughout the QC Case Review system. Instead of scattering values like
'   "QC" (the sheet password) or network paths throughout the code, we
'   define them here once as named constants.
'
' WHY THIS EXISTS:
'   The original codebase had values hardcoded in many places:
'   - The password "QC" appeared dozens of times
'   - Network paths were repeated throughout
'   - Review number prefixes were checked with literals like "50", "51"
'
'   Problems with hardcoded values:
'   - If the password changes, you have to find and update EVERY occurrence
'   - Easy to miss some, leaving inconsistent behavior
'   - Hard to understand what "50" means without context
'
'   Benefits of using this module:
'   - Change a value in ONE place, it updates everywhere
'   - Constants have meaningful names (SHEET_PASSWORD vs "QC")
'   - New maintainers can quickly see all configurable values
'
' HOW TO USE:
'   From any module, just use the constant names directly:
'
'       ws.Protect Password:=SHEET_PASSWORD
'       ' Uses "QC" but you don't have to remember that
'
'       basePath = UNC_PATH_SERVER & UNC_PATH_SHARE
'       ' Builds the full network path
'
' SECTIONS IN THIS MODULE:
'   1. Network Paths      - Server and share locations
'   2. Sheet Protection   - Password for protected worksheets
'   3. Review Prefixes    - Program-identifying prefixes
'   4. Income Multipliers - For frequency conversions
'   5. Code Translations  - BIS code to display value mappings
'   6. Cell References    - Key cells in schedule templates
'
' MAINTENANCE NOTES:
'   - When something changes (new password, new server), update HERE
'   - Test after changes - one typo here affects the whole system
'   - Keep this file backed up - it's critical configuration
'
' CHANGE LOG:
'   2026-01-02  Initial creation - consolidated from scattered hardcoded values
' ============================================================================

Option Explicit


' ============================================================================
' SECTION 1: NETWORK PATHS
' ============================================================================
' PURPOSE:
'   These constants define the network server and share paths where all
'   QC review files are stored. The full path is built by combining
'   these pieces: \\server\share1\share2\data\stat\dqc\
'
' IF THE SERVER CHANGES:
'   Update UNC_PATH_SERVER with the new server name.
'   All code using these constants will automatically use the new path.
'
' NOTES:
'   UNC = Universal Naming Convention (the \\server\share format)
'   We split into parts for flexibility, but also provide the full path
' ============================================================================

' The file server name
Public Const UNC_SERVER As String = "hsedcprapfpp001"

' The full UNC path to the "stat" folder (where DQC lives)
' Some users have their drive mapped here
Public Const UNC_PATH_STAT As String = "\\hsedcprapfpp001\oim\pwimdaubts04\data\stat"

' The full UNC path to the DQC folder directly
' Some users have their drive mapped here instead
Public Const UNC_PATH_DQC As String = "\\hsedcprapfpp001\oim\pwimdaubts04\data\stat\dqc"

' Subfolder names within DQC
Public Const FOLDER_SCHEDULES As String = "Schedules by Examiner Number"
Public Const FOLDER_TEMPLATES As String = "Templates"
Public Const FOLDER_HQ_DATA_ENTRY As String = "HQ - Data Entry"
Public Const FOLDER_CREATE_FO_DB As String = "Create FO Databases"
Public Const FOLDER_FO_DATABASES As String = "FO Databases"


' ============================================================================
' SECTION 2: SHEET PROTECTION
' ============================================================================
' PURPOSE:
'   The schedule worksheets are protected to prevent accidental edits.
'   This password is used to protect and unprotect sheets.
'
' IF THE PASSWORD CHANGES:
'   Update SHEET_PASSWORD and all schedules will work with the new password.
'   NOTE: Existing schedules with the old password will need to be updated.
'
' SECURITY NOTE:
'   Excel sheet protection is NOT secure. It prevents accidents, not
'   determined users. Anyone with basic VBA knowledge can unprotect sheets.
'   Don't store sensitive data relying on this protection.
' ============================================================================

' Password used for all protected worksheets
Public Const SHEET_PASSWORD As String = "QC"


' ============================================================================
' SECTION 3: REVIEW NUMBER PREFIXES
' ============================================================================
' PURPOSE:
'   Review numbers have prefixes that identify which program they belong to.
'   These constants document what each prefix means and are used by the
'   GetProgramFromReviewNumber function in Common_Utils.
'
' HOW REVIEW NUMBERS WORK:
'   - First 1-2 digits = program identifier
'   - Remaining digits = sequential case number
'   - Example: 50012345 = SNAP Positive, case 12345
'
' IF NEW PREFIXES ARE ADDED:
'   1. Add the constant(s) here
'   2. Update GetProgramFromReviewNumber in Common_Utils
'   3. Update GetProgramFolderName in Common_Utils
' ============================================================================

' SNAP Positive review prefixes
Public Const PREFIX_SNAP_POS_50 As String = "50"
Public Const PREFIX_SNAP_POS_51 As String = "51"
Public Const PREFIX_SNAP_POS_55 As String = "55"

' SNAP Negative review prefixes
Public Const PREFIX_SNAP_NEG_60 As String = "60"
Public Const PREFIX_SNAP_NEG_61 As String = "61"
Public Const PREFIX_SNAP_NEG_65 As String = "65"
Public Const PREFIX_SNAP_NEG_66 As String = "66"

' TANF review prefix
Public Const PREFIX_TANF As String = "14"

' GA (General Assistance) review prefix
' Note: GA uses single digit "9" not "9x"
Public Const PREFIX_GA As String = "9"

' MA Positive review prefixes
Public Const PREFIX_MA_POS_20 As String = "20"
Public Const PREFIX_MA_POS_21 As String = "21"

' MA PE (Presumptive Eligibility) review prefix
Public Const PREFIX_MA_PE As String = "24"

' MA Negative review prefixes
Public Const PREFIX_MA_NEG_80 As String = "80"
Public Const PREFIX_MA_NEG_81 As String = "81"
Public Const PREFIX_MA_NEG_82 As String = "82"
Public Const PREFIX_MA_NEG_83 As String = "83"


' ============================================================================
' SECTION 4: INCOME FREQUENCY MULTIPLIERS
' ============================================================================
' PURPOSE:
'   When calculating monthly income, we need to convert from the reported
'   frequency to monthly. These multipliers are used for that conversion.
'
' BIS FREQUENCY CODES (original array indexed 0-9):
'   Index 0 = No income / Invalid -> 0
'   Index 1 = Monthly             -> 1 (already monthly)
'   Index 2 = Weekly              -> 4 (4 weeks/month approx)
'   Index 3 = Bi-Weekly           -> 2 (2 pay periods/month)
'   Index 4 = Semi-Monthly        -> 2 (2 pay periods/month)
'   Index 5 = Monthly (dup)       -> 1
'   Index 6 = Bi-Monthly          -> 0.5 (every 2 months)
'   Index 7 = Quarterly           -> 0.333333 (every 3 months)
'   Index 8 = Semi-Annually       -> 0.166667 (every 6 months)
'   Index 9 = Annually            -> 0.083333 (every 12 months)
'
' EXAMPLE:
'   If someone earns $500 weekly (frequency code 2):
'   Monthly = 500 * GetIncomeMultiplierByIndex(2) = 500 * 4 = $2,000
'
' ORIGINAL ARRAY FROM V1:
'   income_freq = Array(0, 1, 4, 2, 2, 1, 0.5, 0.333333, 0.166667, 0.083333)
' ============================================================================

' Individual multiplier constants for direct access if needed
Public Const INCOME_MULT_NONE As Double = 0
Public Const INCOME_MULT_MONTHLY As Double = 1
Public Const INCOME_MULT_WEEKLY As Double = 4
Public Const INCOME_MULT_BIWEEKLY As Double = 2
Public Const INCOME_MULT_SEMIMONTHLY As Double = 2
Public Const INCOME_MULT_BIMONTHLY As Double = 0.5
Public Const INCOME_MULT_QUARTERLY As Double = 0.333333
Public Const INCOME_MULT_SEMIANNUAL As Double = 0.166667
Public Const INCOME_MULT_ANNUAL As Double = 0.083333

' Note: VBA doesn't support constant arrays, so GetIncomeMultiplierByIndex() provides array access


' ============================================================================
' SECTION 5: STATUS FOLDER NAMES
' ============================================================================
' PURPOSE:
'   When reviews are completed, they're filed into folders based on status.
'   These constants define those folder names.
'
' FOLDER STRUCTURE:
'   .../Examiner Name - #/Program/Review Month XXXX YYYY/
'       Clean/     <- Reviews with no errors
'       Error/     <- Reviews that found errors
'       Drop/      <- Reviews that were dropped
' ============================================================================

Public Const STATUS_FOLDER_CLEAN As String = "Clean"
Public Const STATUS_FOLDER_ERROR As String = "Error"
Public Const STATUS_FOLDER_DROP As String = "Drop"


' ============================================================================
' SECTION 6: BIS FILE CONSTANTS
' ============================================================================
' PURPOSE:
'   Constants related to the BIS (Benefits Information System) delimited
'   files that are used to populate schedules.
'
' BIS FILE STRUCTURE:
'   Tab-delimited text files with:
'   - "Case" worksheet: One row of case-level data
'   - "Individual" worksheet: One row per household member
' ============================================================================

' Row where data starts in BIS files (after header)
Public Const BIS_DATA_START_ROW As Long = 2

' Column positions in BIS Case worksheet (0-indexed for array access)
' These are documented here but actual positions are in data mapping docs
' Add specific column constants as needed for refactoring


' ============================================================================
' SECTION 7: SCHEDULE TEMPLATE CELL REFERENCES
' ============================================================================
' PURPOSE:
'   Key cell addresses in the schedule templates. When templates are updated,
'   these may need to change. Documenting them here makes updates easier.
'
' ORGANIZATION:
'   Grouped by program type, then by purpose within each program.
'
' CELL REFERENCE FORMAT:
'   Constants are named: CELL_{PROGRAM}_{PURPOSE}
'   Example: CELL_TANF_DISPOSITION = "AI10"
' ============================================================================

' --- TANF Schedule Cell References ---
Public Const CELL_TANF_REVIEW_NUMBER As String = "A10"
Public Const CELL_TANF_CASE_NUMBER As String = "I10"
Public Const CELL_TANF_DISPOSITION As String = "AI10"
Public Const CELL_TANF_REVIEW_FINDINGS As String = "AL10"
Public Const CELL_TANF_ERROR_AMOUNT As String = "AO10"
Public Const CELL_TANF_SAMPLE_MONTH As String = "AB10"

' --- SNAP Positive Schedule Cell References ---
Public Const CELL_SNAP_POS_DISPOSITION As String = "C22"
Public Const CELL_SNAP_POS_FINDING_CODE As String = "K22"

' --- SNAP Negative Schedule Cell References ---
Public Const CELL_SNAP_NEG_DISPOSITION As String = "F29"
Public Const CELL_SNAP_NEG_ERROR_ELEMENT As String = "E47"

' --- MA Positive Schedule Cell References ---
Public Const CELL_MA_POS_INIT_ELIG As String = "F16"
Public Const CELL_MA_POS_REVIEW_STATUS As String = "S16"

' --- MA Negative Schedule Cell References ---
Public Const CELL_MA_NEG_DISPOSITION As String = "M56"
Public Const CELL_MA_NEG_HEARING_REQ As String = "C25"
Public Const CELL_MA_NEG_ELIG_REQ As String = "C40"
Public Const CELL_MA_NEG_FIELD_INV As String = "C56"


' ============================================================================
' SECTION 8: TEMPLATE PATHS
' ============================================================================
' PURPOSE:
'   Paths to Word/Excel template files used for generating memos, letters,
'   and other documents. All paths are relative to the DQC root folder.
'
' USAGE:
'   fullPath = GetDQCDriveLetter() & TEMPLATE_FINDINGS_MEMO_SNAP
' ============================================================================

' --- Findings Memo Templates ---
Public Const TEMPLATE_FINDINGS_MEMO_SNAP_POS As String = "Finding Memo\SNAP Positive QC Finding Memo.xlsm"
Public Const TEMPLATE_FINDINGS_MEMO_SNAP_NEG As String = "Finding Memo\SNAP Negative QC Finding Memo.xlsm"
Public Const TEMPLATE_FINDINGS_MEMO_TANF As String = "Finding Memo\TANF QC Finding Memo.xlsm"
Public Const TEMPLATE_FINDINGS_MEMO_GA As String = "Finding Memo\GA QC Finding Memo.xlsm"
Public Const TEMPLATE_FINDINGS_MEMO_MA_POS As String = "Finding Memo\MA Positive QC Finding Memo.xlsm"
Public Const TEMPLATE_FINDINGS_MEMO_MA_NEG As String = "Finding Memo\MA Negative QC Finding Memo.xlsm"
Public Const TEMPLATE_TIMELINESS_MEMO As String = "Finding Memo\SNAP Timeliness QC Finding Memo.xlsm"

' --- Data Source Templates ---
Public Const TEMPLATE_FINDING_MEMO_DATA_SOURCE As String = "Finding Memo\Finding Memo Data Source.xlsx"
Public Const TEMPLATE_COURTESY_COPY_LIST As String = "Finding Memo\Courtesy Copy for Memos.xlsx"

' --- Appointment Letter Templates ---
Public Const TEMPLATE_CAO_APPOINTMENT As String = "Templates\CAO Appointment Letter.docx"
Public Const TEMPLATE_CAO_APPOINTMENT_SPANISH As String = "Templates\CAO Appointment Letter Spanish.docx"
Public Const TEMPLATE_TELEPHONE_APPOINTMENT As String = "Templates\Telephone Appointment Letter.docx"
Public Const TEMPLATE_TELEPHONE_APPOINTMENT_SPANISH As String = "Templates\Telephone Appointment Letter Spanish.docx"
Public Const TEMPLATE_MA_APPOINTMENT As String = "Templates\MA Appointment Letter.docx"
Public Const TEMPLATE_MA_APPOINTMENT_SPANISH As String = "Templates\MA Appointment Letter Spanish.docx"

' --- Information and Notification Memos ---
Public Const TEMPLATE_INFO_MEMO As String = "Templates\Information Memo.docx"
Public Const TEMPLATE_PENDING_LETTER As String = "Templates\Pending Letter.docx"
Public Const TEMPLATE_PENDING_LETTER_SPANISH As String = "Templates\Pending Letter Spanish.docx"
Public Const TEMPLATE_CASE_SUMMARY As String = "Templates\Case Summary Template.docx"
Public Const TEMPLATE_DROP_WORKSHEET As String = "Templates\Drop Worksheet.xlsx"

' --- QC14 Series Forms ---
Public Const TEMPLATE_QC14_COOP As String = "Templates\QC14 Coop Memo.docx"
Public Const TEMPLATE_QC14_CAO_REQUEST As String = "Templates\QC14 CAO Request.docx"
Public Const TEMPLATE_QC14F As String = "Templates\QC14F.docx"
Public Const TEMPLATE_QC14R As String = "Templates\QC14R.docx"
Public Const TEMPLATE_QC14C As String = "Templates\QC14C.docx"
Public Const TEMPLATE_QC15 As String = "Templates\QC15.docx"

' --- Other Templates ---
Public Const TEMPLATE_LEP As String = "Templates\LEP.docx"
Public Const TEMPLATE_POST_OFFICE As String = "Templates\Post Office Memo.docx"
Public Const TEMPLATE_THRESHOLD As String = "Templates\Error Under Threshold Memo.docx"
Public Const TEMPLATE_POTENTIAL_ERROR As String = "Templates\Potential Error Call Memo.docx"
Public Const TEMPLATE_AMR As String = "Templates\AMR Memo.docx"
Public Const TEMPLATE_CRIMINAL_HISTORY As String = "Templates\Criminal History.docx"
Public Const TEMPLATE_SELF_EMPLOYMENT As String = "Templates\Self Employment.docx"
Public Const TEMPLATE_SELF_EMPLOYMENT_DETAILED As String = "Templates\Self Employment Detailed.docx"
Public Const TEMPLATE_HOUSEHOLD_COMP As String = "Templates\Household Composition.docx"
Public Const TEMPLATE_SAVE_DEFICIENCY As String = "Templates\SAVE Deficiency Memo.docx"
Public Const TEMPLATE_COMMUNITY_SPOUSE As String = "Templates\Community Spouse Questionnaire.docx"
Public Const TEMPLATE_PA472_PENDING As String = "Templates\PA472 Pending.docx"
Public Const TEMPLATE_PA76 As String = "Templates\PA76.docx"
Public Const TEMPLATE_PA78 As String = "Templates\PA78.docx"


' ============================================================================
' SECTION 9: MAIL MERGE FIELD NAMES
' ============================================================================
' PURPOSE:
'   Standard placeholder names used in Word templates for mail merge.
'   Using constants prevents typos and makes it easy to update field names.
' ============================================================================

' --- Client Information Fields ---
Public Const FIELD_CLIENT_NAME As String = "<<CLIENT_NAME>>"
Public Const FIELD_CLIENT_ADDRESS As String = "<<CLIENT_ADDRESS>>"
Public Const FIELD_CLIENT_CITY_STATE_ZIP As String = "<<CLIENT_CITY_STATE_ZIP>>"

' --- Case Information Fields ---
Public Const FIELD_CASE_NUMBER As String = "<<CASE_NUMBER>>"
Public Const FIELD_REVIEW_NUMBER As String = "<<REVIEW_NUMBER>>"
Public Const FIELD_SAMPLE_MONTH As String = "<<SAMPLE_MONTH>>"
Public Const FIELD_COUNTY_NAME As String = "<<COUNTY_NAME>>"
Public Const FIELD_COUNTY_NUMBER As String = "<<COUNTY_NUMBER>>"
Public Const FIELD_DISTRICT_NUMBER As String = "<<DISTRICT_NUMBER>>"
Public Const FIELD_PROCESSING_CENTER As String = "<<PROCESSING_CENTER>>"

' --- Error/Finding Fields ---
Public Const FIELD_ERROR_TYPE As String = "<<ERROR_TYPE>>"
Public Const FIELD_ERROR_AMOUNT As String = "<<ERROR_AMOUNT>>"
Public Const FIELD_BENEFIT_AMOUNT As String = "<<BENEFIT_AMOUNT>>"
Public Const FIELD_QC_BENEFIT_AMOUNT As String = "<<QC_BENEFIT_AMOUNT>>"

' --- Personnel Fields ---
Public Const FIELD_EXAMINER_NAME As String = "<<EXAMINER_NAME>>"
Public Const FIELD_SUPERVISOR_NAME As String = "<<SUPERVISOR_NAME>>"
Public Const FIELD_PROGRAM_MANAGER As String = "<<PROGRAM_MANAGER>>"
Public Const FIELD_CAO_DIRECTOR As String = "<<CAO_DIRECTOR>>"

' --- Date/Time Fields ---
Public Const FIELD_TODAY_DATE As String = "<<TODAY_DATE>>"
Public Const FIELD_APPOINTMENT_DATE As String = "<<APPOINTMENT_DATE>>"
Public Const FIELD_APPOINTMENT_TIME As String = "<<APPOINTMENT_TIME>>"


' ============================================================================
' SECTION 10: PROGRAM-SPECIFIC CELL MAPPINGS
' ============================================================================
' PURPOSE:
'   Defines which cells contain key data in each program's schedule template.
'   Used by extraction functions to batch-read data efficiently.
'
' NOTE:
'   Column numbers are 1-based (A=1, B=2, ..., AJ=36, etc.)
' ============================================================================

' --- SNAP Positive Data Cells ---
Public Const SNAP_POS_ROW_CLIENT_NAME As Long = 4
Public Const SNAP_POS_COL_CLIENT_NAME As Long = 2       ' B4
Public Const SNAP_POS_ROW_REVIEW_NUM As Long = 18
Public Const SNAP_POS_COL_REVIEW_NUM As Long = 1        ' A18
Public Const SNAP_POS_ROW_CASE_NUM As Long = 18
Public Const SNAP_POS_COL_CASE_NUM As Long = 9          ' I18
Public Const SNAP_POS_ROW_EXAMINER As Long = 5
Public Const SNAP_POS_COL_EXAMINER_1 As Long = 36       ' AJ5
Public Const SNAP_POS_COL_EXAMINER_2 As Long = 37       ' AK5
Public Const SNAP_POS_ROW_DISPOSITION As Long = 22
Public Const SNAP_POS_COL_DISPOSITION As Long = 3       ' C22
Public Const SNAP_POS_ROW_FINDING As Long = 22
Public Const SNAP_POS_COL_FINDING As Long = 11          ' K22

' --- TANF/GA Data Cells ---
Public Const TANF_ROW_REVIEW_NUM As Long = 10
Public Const TANF_COL_REVIEW_NUM As Long = 1            ' A10
Public Const TANF_ROW_CASE_NUM As Long = 10
Public Const TANF_COL_CASE_NUM As Long = 9              ' I10
Public Const TANF_ROW_EXAMINER As Long = 3
Public Const TANF_COL_EXAMINER_1 As Long = 41           ' AO3
Public Const TANF_COL_EXAMINER_2 As Long = 42           ' AP3
Public Const TANF_ROW_DISPOSITION As Long = 10
Public Const TANF_COL_DISPOSITION As Long = 35          ' AI10
Public Const TANF_ROW_FINDING As Long = 10
Public Const TANF_COL_FINDING As Long = 38              ' AL10
Public Const TANF_ROW_ERROR_AMOUNT As Long = 10
Public Const TANF_COL_ERROR_AMOUNT As Long = 41         ' AO10


' ============================================================================
' SECTION 11: EMAIL CONFIGURATION
' ============================================================================
' PURPOSE:
'   Constants for the email notifications sent when reviews are returned
'   to examiners or completed.
' ============================================================================

' Email subject line templates
Public Const EMAIL_SUBJECT_RETURN As String = "QC Review Returned - "
Public Const EMAIL_SUBJECT_COMPLETE As String = "QC Review Complete - "

' Email importance levels
Public Const EMAIL_IMPORTANCE_HIGH As Long = 2
Public Const EMAIL_IMPORTANCE_NORMAL As Long = 1


' ============================================================================
' SECTION 9: HELPER FUNCTIONS FOR CONFIGURATION
' ============================================================================
' PURPOSE:
'   Some configuration values can't be constants (like arrays). These
'   functions provide access to those values.
' ============================================================================

' ----------------------------------------------------------------------------
' Function: GetIncomeMultiplierByIndex
' ----------------------------------------------------------------------------
' PURPOSE:
'   Returns the income frequency multiplier for the given BIS index.
'   This matches the original income_freq array used in V1:
'   income_freq = Array(0, 1, 4, 2, 2, 1, 0.5, 0.333333, 0.166667, 0.083333)
'
' PARAMETERS:
'   freqIndex - The BIS frequency code index (0-9)
'
' RETURNS:
'   Double - The multiplier to convert to monthly
'            Returns 1 if index is invalid (assumes monthly)
'
' EXAMPLE:
'   monthlyAmount = weeklyAmount * GetIncomeMultiplierByIndex(2)  ' Weekly
' ----------------------------------------------------------------------------
Public Function GetIncomeMultiplierByIndex(ByVal freqIndex As Long) As Double
    Select Case freqIndex
        Case 0
            GetIncomeMultiplierByIndex = INCOME_MULT_NONE
        Case 1
            GetIncomeMultiplierByIndex = INCOME_MULT_MONTHLY
        Case 2
            GetIncomeMultiplierByIndex = INCOME_MULT_WEEKLY
        Case 3
            GetIncomeMultiplierByIndex = INCOME_MULT_BIWEEKLY
        Case 4
            GetIncomeMultiplierByIndex = INCOME_MULT_SEMIMONTHLY
        Case 5
            GetIncomeMultiplierByIndex = INCOME_MULT_MONTHLY
        Case 6
            GetIncomeMultiplierByIndex = INCOME_MULT_BIMONTHLY
        Case 7
            GetIncomeMultiplierByIndex = INCOME_MULT_QUARTERLY
        Case 8
            GetIncomeMultiplierByIndex = INCOME_MULT_SEMIANNUAL
        Case 9
            GetIncomeMultiplierByIndex = INCOME_MULT_ANNUAL
        Case Else
            ' Default to monthly if unknown code
            GetIncomeMultiplierByIndex = 1
    End Select
End Function

' ----------------------------------------------------------------------------
' Function: GetIncomeMultiplierArray
' ----------------------------------------------------------------------------
' PURPOSE:
'   Returns the full income multiplier array as a Variant.
'   Use this when you need array-style access for performance.
'
' RETURNS:
'   Variant - Array of multipliers indexed 0-9
'
' EXAMPLE:
'   Dim incFreq As Variant
'   incFreq = GetIncomeMultiplierArray()
'   monthlyAmt = weeklyAmt * incFreq(2)
' ----------------------------------------------------------------------------
Public Function GetIncomeMultiplierArray() As Variant
    GetIncomeMultiplierArray = Array(0, 1, 4, 2, 2, 1, 0.5, 0.333333, 0.166667, 0.083333)
End Function

' ----------------------------------------------------------------------------
' Function: BuildExaminerPath
' ----------------------------------------------------------------------------
' PURPOSE:
'   Builds the full path to an examiner's program folder.
'
' PARAMETERS:
'   dqcDrive     - The DQC drive letter (from GetDQCDriveLetter)
'   examinerName - The examiner's name
'   examinerNum  - The examiner's number (will be formatted)
'   programFolder- The program folder name (from GetProgramFolderName)
'
' RETURNS:
'   String - Full path like "E:\DQC\Schedules by Examiner Number\John Smith - 5\TANF\"
'
' EXAMPLE:
'   path = BuildExaminerPath("E:\DQC\", "John Smith", "5", "TANF")
' ----------------------------------------------------------------------------
Public Function BuildExaminerPath(ByVal dqcDrive As String, _
                                   ByVal examinerName As String, _
                                   ByVal examinerNum As String, _
                                   ByVal programFolder As String) As String
    
    BuildExaminerPath = dqcDrive & _
                        FOLDER_SCHEDULES & "\" & _
                        examinerName & " - " & examinerNum & "\" & _
                        programFolder & "\"
End Function

' ----------------------------------------------------------------------------
' Function: BuildReviewMonthPath
' ----------------------------------------------------------------------------
' PURPOSE:
'   Builds the "Review Month" folder name from month and year.
'
' PARAMETERS:
'   monthName - Full month name ("January", "February", etc.)
'   year      - 4-digit year
'
' RETURNS:
'   String - Folder name like "Review Month January 2024"
'
' EXAMPLE:
'   folderName = BuildReviewMonthPath("January", "2024")
'   ' Returns "Review Month January 2024"
' ----------------------------------------------------------------------------
Public Function BuildReviewMonthPath(ByVal monthName As String, _
                                      ByVal year As String) As String
    BuildReviewMonthPath = "Review Month " & monthName & " " & year
End Function

' ----------------------------------------------------------------------------
' Function: GetDatabasePath
' ----------------------------------------------------------------------------
' PURPOSE:
'   Returns the path where blank database templates are stored.
'
' PARAMETERS:
'   dqcDrive - The DQC drive letter
'
' RETURNS:
'   String - Path to the FO Databases folder
' ----------------------------------------------------------------------------
Public Function GetDatabasePath(ByVal dqcDrive As String) As String
    GetDatabasePath = dqcDrive & _
                      FOLDER_HQ_DATA_ENTRY & "\" & _
                      FOLDER_CREATE_FO_DB & "\" & _
                      FOLDER_FO_DATABASES & "\"
End Function

' ----------------------------------------------------------------------------
' Function: GetTemplatePath
' ----------------------------------------------------------------------------
' PURPOSE:
'   Builds the full path to a template file.
'
' PARAMETERS:
'   dqcDrive     - The DQC drive letter (from GetDQCDriveLetter)
'   templateName - The template constant (e.g., TEMPLATE_FINDINGS_MEMO_TANF)
'
' RETURNS:
'   String - Full path to the template file
'
' EXAMPLE:
'   path = GetTemplatePath("E:\DQC\", TEMPLATE_FINDINGS_MEMO_TANF)
' ----------------------------------------------------------------------------
Public Function GetTemplatePath(ByVal dqcDrive As String, _
                                 ByVal templateName As String) As String
    GetTemplatePath = dqcDrive & templateName
End Function

' ----------------------------------------------------------------------------
' Function: GetFindingsMemoTemplate
' ----------------------------------------------------------------------------
' PURPOSE:
'   Returns the appropriate findings memo template path for a given program.
'
' PARAMETERS:
'   dqcDrive    - The DQC drive letter
'   programType - The ProgramType enum value
'
' RETURNS:
'   String - Full path to the appropriate findings memo template
' ----------------------------------------------------------------------------
Public Function GetFindingsMemoTemplate(ByVal dqcDrive As String, _
                                         ByVal programType As ProgramType) As String
    Dim templateName As String
    
    Select Case programType
        Case PROG_SNAP_POS
            templateName = TEMPLATE_FINDINGS_MEMO_SNAP_POS
        Case PROG_SNAP_NEG
            templateName = TEMPLATE_FINDINGS_MEMO_SNAP_NEG
        Case PROG_TANF
            templateName = TEMPLATE_FINDINGS_MEMO_TANF
        Case PROG_GA
            templateName = TEMPLATE_FINDINGS_MEMO_GA
        Case PROG_MA_POS
            templateName = TEMPLATE_FINDINGS_MEMO_MA_POS
        Case PROG_MA_NEG, PROG_MA_PE
            templateName = TEMPLATE_FINDINGS_MEMO_MA_NEG
        Case Else
            templateName = ""
    End Select
    
    If templateName <> "" Then
        GetFindingsMemoTemplate = dqcDrive & templateName
    Else
        GetFindingsMemoTemplate = ""
    End If
End Function


