' ============================================================================
' TANF Edit Check V2 - Fully Optimized Version
' ============================================================================
' REQUIREMENTS:
'   - Reference: Microsoft ActiveX Data Objects 6.1 Library (or 2.8+)
'   - Access Database Engine installed for ACE OLEDB 12.0 provider
' ============================================================================
Option Explicit

' Output row counters
Private rswr As Long, qcwr As Long, plwr As Long, hiwr As Long, efwr As Long

' Pre-allocated output arrays (sized for max expected records)
Private rsData() As Variant   ' Review Summary
Private qcData() As Variant   ' QC Case Info
Private plData() As Variant   ' Person Level Info
Private hiData() As Variant   ' Household Income
Private efData() As Variant   ' Error Findings

' Column counts for each table
Private Const RS_COLS As Long = 15
Private Const QC_COLS As Long = 23
Private Const PL_COLS As Long = 18
Private Const HI_COLS As Long = 5
Private Const EF_COLS As Long = 12

' ==========================================================================
' SOURCE TEMPLATE LAYOUT CONSTANTS
' Update these if the source workbook template structure changes
' ==========================================================================
' Max rows per case for child tables (used for array pre-allocation)
Private Const MAX_PERSONS_PER_CASE As Long = 8    ' Rows 30-44 step 2 = 8 persons max
Private Const MAX_INCOME_PER_CASE As Long = 16    ' 4 rows × 4 income types = 16 max
Private Const MAX_ERRORS_PER_CASE As Long = 4     ' Rows 61-67 step 2 = 4 errors max

' Source data range - covers all used rows (highest data row is 85 for AB85 Renewal Type)
' Rows used: 3, 10, 16, 20, 24, 30-44, 50-56, 61-67, 85
Private Const SRC_CACHE_RANGE As String = "A1:AQ85"

' Source workbook data cache
Private srcCache As Variant
Private revidval As Long

Sub Find_Write_Database_Files_V2()
    On Error GoTo ErrorHandler
    
    ' ========================================================================
    ' 1. PERFORMANCE OPTIMIZATION - Disable all Excel overhead
    ' ========================================================================
    With Application
        .ScreenUpdating = False
        .DisplayStatusBar = True
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With
    
    ' Variable declarations
    Dim i As Long, j As Long, n As Long
    Dim maxrow As Long, maxrowex As Long
    Dim program As String, exname As String, exnumstr As String
    Dim reviewtxt As String, disp_code As Variant
    Dim pathdir As String, BasePath As String, CaseFolderPath As String
    Dim CaseFolderName As String, FinalFilePath As String, FileNameFound As String
    Dim monthstr As String, mName As String, yStr As String
    Dim sPath As String, SrceFile As String, exceloutfile As String, databasename As String
    Dim DLetter As String
    Dim filenum As Integer
    Dim processedCount As Long
    
    ' Workbook/Worksheet objects
    Dim thissht As Worksheet
    Dim inWB As Workbook, outWB As Workbook
    Dim inWS As Worksheet
    
    ' Network objects
    Dim WshNetwork As Object, oDrives As Object
    
    ' ADO objects (Early Binding for speed - requires reference)
    Dim cnt As ADODB.Connection
    Dim rs As ADODB.Recordset
    
    ' Input data arrays
    Dim vInput As Variant, vExLookup As Variant
    Dim currentExNum As String
    
    Set thissht = ActiveSheet
    sPath = ActiveWorkbook.Path
    
    ' ========================================================================
    ' 2. NETWORK PATH VALIDATION
    ' ========================================================================
    Set WshNetwork = CreateObject("WScript.Network")
    Set oDrives = WshNetwork.EnumNetworkDrives
    
    DLetter = ""
    For i = 0 To oDrives.Count - 1 Step 2
        Select Case LCase(oDrives.Item(i + 1))
            Case "\\hsedcprapfpp001\oim\pwimdaubts04\data\stat"
                DLetter = oDrives.Item(i) & "\DQC\"
                Exit For
            Case "\\hsedcprapfpp001\oim\pwimdaubts04\data\stat\dqc"
                DLetter = oDrives.Item(i) & "\"
                Exit For
        End Select
    Next i
    
    If DLetter = "" Then
        Err.Raise 9999, "Find_Write_Database_Files_V2", _
            "Network Drive to Examiner Files NOT found." & vbCrLf & "Contact Valerie or Wes."
    End If
    
    pathdir = DLetter & "Schedules by Examiner Number\"
    If Dir(pathdir, vbDirectory) = "" Then
        Err.Raise 9999, "Find_Write_Database_Files_V2", _
            "Examiner Directory not found: " & pathdir
    End If
    
    ' ========================================================================
    ' 3. LOAD ALL INPUT DATA INTO MEMORY (Single read operations)
    ' ========================================================================
    maxrow = thissht.Range("E" & thissht.Rows.Count).End(xlUp).Row
    maxrowex = thissht.Range("L" & thissht.Rows.Count).End(xlUp).Row
    
    ' Load review data: E=ReviewNum, F=Month, G=ExaminerNum
    vInput = thissht.Range("E1:G" & maxrow).Value
    ' Load examiner lookup: K=Name, L=Number
    vExLookup = thissht.Range("K1:L" & maxrowex).Value
    
    ' ========================================================================
    ' 4. PRE-ALLOCATE OUTPUT ARRAYS (Avoids incremental resizing)
    ' ========================================================================
    Dim maxCases As Long
    maxCases = maxrow - 1  ' Estimate max records
    
    ReDim rsData(1 To maxCases, 1 To RS_COLS)
    ReDim qcData(1 To maxCases, 1 To QC_COLS)
    ReDim plData(1 To maxCases * MAX_PERSONS_PER_CASE, 1 To PL_COLS)
    ReDim hiData(1 To maxCases * MAX_INCOME_PER_CASE, 1 To HI_COLS)
    ReDim efData(1 To maxCases * MAX_ERRORS_PER_CASE, 1 To EF_COLS)
    
    ' Initialize counters (arrays are 1-based, row 1 = first data row)
    rswr = 0: qcwr = 0: plwr = 0: hiwr = 0: efwr = 0
    processedCount = 0
    
    ' ========================================================================
    ' 5. MAIN PROCESSING LOOP
    ' ========================================================================
    For i = 2 To maxrow
        ' Update status every 10 records (reduces DoEvents overhead)
        If i Mod 10 = 0 Or i = maxrow Then
            Application.StatusBar = "Processing " & (i - 1) & "/" & (maxrow - 1) & _
                " (" & Format((i - 1) / (maxrow - 1), "0%") & ")"
            DoEvents
        End If
        
        ' ---------------------------------------------------------------------
        ' A. Parse and validate review number
        ' ---------------------------------------------------------------------
        reviewtxt = Trim(CStr(vInput(i, 1)))
        If Left(reviewtxt, 1) = "0" Then reviewtxt = Mid(reviewtxt, 2)
        
        ' Only process TANF reviews (start with "1")
        If Left(reviewtxt, 1) <> "1" Then GoTo NextIteration
        program = "TANF"
        
        ' ---------------------------------------------------------------------
        ' B. Find examiner name (array lookup - no Excel calls)
        ' ---------------------------------------------------------------------
        exname = ""
        currentExNum = Trim(CStr(vInput(i, 3)))
        
        For j = 2 To maxrowex
            If Trim(CStr(vExLookup(j, 2))) = currentExNum Then
                exname = Trim(CStr(vExLookup(j, 1)))
                Exit For
            End If
        Next j
        
        If exname = "" Then GoTo NextIteration
        
        ' Format examiner number (remove leading zero)
        exnumstr = Format(currentExNum, "00")
        If Left(exnumstr, 1) = "0" Then exnumstr = Right(exnumstr, 1)
        
        ' ---------------------------------------------------------------------
        ' C. Parse sample month
        ' ---------------------------------------------------------------------
        monthstr = Trim(CStr(vInput(i, 2)))
        If Len(monthstr) < 6 Then GoTo NextIteration
        
        yStr = Left(monthstr, 4)
        mName = MonthName(Val(Right(monthstr, 2)))
        
        ' ---------------------------------------------------------------------
        ' D. Build path and find workbook
        ' ---------------------------------------------------------------------
        BasePath = pathdir & exname & " - " & exnumstr & "\" & program & "\" & _
                   "Review Month " & mName & " " & yStr & "\"
        
        ' Find case folder (e.g., "12345 - Smith")
        CaseFolderName = Dir(BasePath & reviewtxt & " - *", vbDirectory)
        
        If CaseFolderName = "" Then GoTo NextIteration
        If (GetAttr(BasePath & CaseFolderName) And vbDirectory) <> vbDirectory Then GoTo NextIteration
        
        CaseFolderPath = BasePath & CaseFolderName & "\"
        FileNameFound = Dir(CaseFolderPath & "Review Number " & reviewtxt & "*.xls*")
        
        If FileNameFound = "" Then GoTo NextIteration
        
        FinalFilePath = CaseFolderPath & FileNameFound
        
        ' ---------------------------------------------------------------------
        ' E. Open and process workbook
        ' ---------------------------------------------------------------------
        Set inWB = Workbooks.Open(Filename:=FinalFilePath, UpdateLinks:=0, ReadOnly:=True)
        
        ' Try to get the review sheet
        On Error Resume Next
        Set inWS = inWB.Sheets(reviewtxt)
        If Err.Number <> 0 Then
            Err.Clear
            inWB.Close False
            Set inWB = Nothing
            On Error GoTo ErrorHandler
            GoTo NextIteration
        End If
        On Error GoTo ErrorHandler
        
        ' ---------------------------------------------------------------------
        ' F. BATCH READ: Load all source data into cache array (single read)
        ' ---------------------------------------------------------------------
        srcCache = inWS.Range(SRC_CACHE_RANGE).Value
        
        disp_code = srcCache(10, 35)  ' AI10 = Row 10, Col 35
        revidval = i - 1
        
        ' ---------------------------------------------------------------------
        ' G. Extract data to output arrays
        ' ---------------------------------------------------------------------
        Call ExtractRevSum(inWB)
        
        If disp_code = 1 Then
            Call ExtractQCInfo
            Call ExtractPLInfo
            Call ExtractHHInc
            Call ExtractErrFind
        End If
        
        inWB.Close False
        Set inWB = Nothing
        processedCount = processedCount + 1
        
NextIteration:
    Next i
    
    ' ========================================================================
    ' 6. WRITE OUTPUT ARRAYS TO EXCEL (Single write per table)
    ' ========================================================================
    Application.StatusBar = "Writing to Excel template..."
    DoEvents
    
    ' Prepare output file
    SrceFile = sPath & "\FO Databases\TANF_Template.xlsx"
    exceloutfile = sPath & "\TANF Database Input " & Format(Date, "mm-dd-yyyy") & ".xlsx"
    If Dir(exceloutfile) <> "" Then Kill exceloutfile
    FileCopy SrceFile, exceloutfile
    Set outWB = Workbooks.Open(exceloutfile)
    
    ' Write arrays to worksheets (single range write per table)
    If rswr > 0 Then
        outWB.Sheets("Review_Summary_dtl").Range("A2").Resize(rswr, RS_COLS).Value = _
            TrimArray(rsData, rswr, RS_COLS)
    End If
    
    If qcwr > 0 Then
        outWB.Sheets("QC_Case_Info_dtl").Range("A2").Resize(qcwr, QC_COLS).Value = _
            TrimArray(qcData, qcwr, QC_COLS)
    End If
    
    If plwr > 0 Then
        outWB.Sheets("Person_Level_Info_dtl").Range("A2").Resize(plwr, PL_COLS).Value = _
            TrimArray(plData, plwr, PL_COLS)
    End If
    
    If hiwr > 0 Then
        outWB.Sheets("Household_Income_dtl").Range("A2").Resize(hiwr, HI_COLS).Value = _
            TrimArray(hiData, hiwr, HI_COLS)
    End If
    
    If efwr > 0 Then
        outWB.Sheets("Error_Findings_dtl").Range("A2").Resize(efwr, EF_COLS).Value = _
            TrimArray(efData, efwr, EF_COLS)
    End If
    
    outWB.Save
    
    ' ========================================================================
    ' 7. DATABASE TRANSFER WITH TRANSACTION
    ' ========================================================================
    Application.StatusBar = "Transferring to Access database..."
    DoEvents
    
    ' Prepare database file
    SrceFile = sPath & "\FO Databases\TANF_Blank.mdb"
    databasename = sPath & "\TANF1 " & Format(Date, "mm-dd-yyyy") & ".mdb"
    filenum = 1
    Do Until Dir(databasename) = ""
        filenum = filenum + 1
        databasename = sPath & "\TANF" & filenum & " " & Format(Date, "mmddyyyy") & ".mdb"
    Loop
    FileCopy SrceFile, databasename
    
    ' Open connection with transaction
    Set cnt = New ADODB.Connection
    cnt.Open "Provider=Microsoft.Ace.OLEDB.12.0;Data Source=" & databasename & ";"
    cnt.BeginTrans  ' Start transaction for batch performance
    
    On Error GoTo DBError
    
    ' Transfer each table
    Call TransferTableToAccess(cnt, outWB.Sheets("Review_Summary_dtl"), "Review_Summary_dtl", rswr)
    Call TransferTableToAccess(cnt, outWB.Sheets("QC_Case_Info_dtl"), "QC_Case_Info_dtl", qcwr)
    Call TransferTableToAccess(cnt, outWB.Sheets("Person_Level_Info_dtl"), "Person_Level_Info_dtl", plwr)
    Call TransferTableToAccess(cnt, outWB.Sheets("Household_Income_dtl"), "Household_Income_dtl", hiwr)
    Call TransferTableToAccess(cnt, outWB.Sheets("Error_Findings_dtl"), "Error_Findings_dtl", efwr)
    
    cnt.CommitTrans  ' Commit all changes at once
    
    On Error GoTo ErrorHandler
    
    ' ========================================================================
    ' 8. CLEANUP
    ' ========================================================================
    cnt.Close
    Set cnt = Nothing
    
    outWB.Close True
    Set outWB = Nothing
    
    ' Delete temporary Excel file
    If Dir(exceloutfile) <> "" Then Kill exceloutfile
    
    ' Clear arrays from memory
    Erase rsData, qcData, plData, hiData, efData
    
    MsgBox "Processing Complete!" & vbCrLf & vbCrLf & _
           "Cases Processed: " & processedCount & vbCrLf & _
           "Database: " & databasename, vbInformation, "TANF Edit Check V2"

CleanExit:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

DBError:
    ' Rollback ALL changes - no partial data
    cnt.RollbackTrans
    MsgBox "DATABASE WRITE FAILED - NO DATA SAVED" & vbCrLf & vbCrLf & _
           "Error: " & Err.Description & vbCrLf & _
           "Error #: " & Err.Number & vbCrLf & vbCrLf & _
           "The transaction has been rolled back." & vbCrLf & _
           "Please review the source data and retry.", vbCritical
    Resume CleanExit

ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & _
           "Source: " & Err.Source, vbCritical, "TANF Edit Check V2 - Error"
    If Not inWB Is Nothing Then inWB.Close False
    If Not outWB Is Nothing Then outWB.Close False
    If Not cnt Is Nothing Then
        If cnt.State = adStateOpen Then cnt.Close
    End If
    Resume CleanExit
End Sub

' ============================================================================
' DATA EXTRACTION SUBROUTINES (All use srcCache array - no Excel calls)
' ============================================================================

Private Sub ExtractRevSum(inWB As Workbook)
    rswr = rswr + 1
    
    rsData(rswr, 1) = revidval                          ' ReviewID
    rsData(rswr, 2) = srcCache(10, 1)                   ' A10 - Review Number
    rsData(rswr, 3) = srcCache(10, 9)                   ' I10 - Case Number
    rsData(rswr, 4) = srcCache(10, 17)                  ' Q10 - Category
    rsData(rswr, 5) = srcCache(10, 19)                  ' S10 - Grant Group
    
    ' Sample Month (AB10 = col 28) - parse MMYYYY format
    Dim sm As Variant: sm = srcCache(10, 28)
    If IsNumeric(sm) And Len(CStr(sm)) >= 6 Then
        Dim smStr As String: smStr = Format(sm, "000000")
        rsData(rswr, 6) = DateSerial(Val(Right(smStr, 4)), Val(Left(smStr, 2)), 1)
    End If
    
    ' Error Amount (only for disposition = 1)
    If srcCache(10, 35) = 1 Then  ' AI10
        If IsNumeric(srcCache(10, 41)) Then  ' AO10
            rsData(rswr, 7) = Round(Val(srcCache(10, 41)) + 0.001, 0)
        End If
    End If
    
    rsData(rswr, 8) = CleanValue(srcCache(10, 38), "B")   ' AL10 - Review Findings
    rsData(rswr, 9) = CleanValue(srcCache(10, 25), "BB")  ' Y10 - District Code
    rsData(rswr, 10) = srcCache(10, 21)                   ' U10 - CAO
    rsData(rswr, 11) = srcCache(10, 35)                   ' AI10 - Disposition Code
    
    ' Run Date from TANF Workbook sheet
    On Error Resume Next
    Dim twSheet As Worksheet
    Set twSheet = inWB.Sheets("TANF Workbook")
    If Not twSheet Is Nothing Then
        rsData(rswr, 12) = twSheet.Range("G33").Value
    End If
    On Error GoTo 0
    
    rsData(rswr, 13) = CStr(srcCache(3, 41)) & CStr(srcCache(3, 42))  ' AO3 & AP3 - Examiner
    rsData(rswr, 15) = CleanValue(srcCache(85, 28), "B")              ' AB85 - Renewal Type
End Sub

Private Sub ExtractQCInfo()
    qcwr = qcwr + 1
    
    qcData(qcwr, 1) = revidval
    qcData(qcwr, 2) = srcCache(20, 22)   ' V20 - Unborn child
    qcData(qcwr, 3) = srcCache(20, 25)   ' Y20 - Shelter Arrangement
    qcData(qcwr, 4) = srcCache(16, 10)   ' J16 - Prior Assistance
    qcData(qcwr, 5) = srcCache(20, 15)   ' O20 - Reason Protective Pay
    qcData(qcwr, 6) = srcCache(16, 21)   ' U16 - Action Type
    qcData(qcwr, 7) = srcCache(16, 1)    ' A16 - Most recent opening
    qcData(qcwr, 8) = srcCache(16, 12)   ' L16 - Most recent action
    qcData(qcwr, 9) = Val(srcCache(16, 23))  ' W16 - Number of case members
    qcData(qcwr, 10) = srcCache(16, 26)  ' Z16 - Liquid assets
    qcData(qcwr, 11) = srcCache(16, 31)  ' AE16 - Real property
    qcData(qcwr, 12) = srcCache(16, 36)  ' AJ16 - Countable Vehicle
    qcData(qcwr, 13) = srcCache(16, 41)  ' AO16 - Non-Liquid Assets
    qcData(qcwr, 14) = srcCache(20, 3)   ' C20 - Monthly Payments
    qcData(qcwr, 15) = srcCache(20, 9)   ' I20 - Sample Month Payments
    qcData(qcwr, 16) = srcCache(20, 17)  ' Q20 - Sanction Amount
    qcData(qcwr, 17) = srcCache(20, 28)  ' AB20 - Gross Income
    qcData(qcwr, 18) = srcCache(20, 34)  ' AH20 - Income Disregard
    qcData(qcwr, 19) = srcCache(20, 40)  ' AN20 - Net Income
    qcData(qcwr, 20) = srcCache(24, 2)   ' B24 - FS Allotment
    qcData(qcwr, 21) = srcCache(24, 21)  ' U24 - Over Payment Recoupment
    qcData(qcwr, 22) = srcCache(20, 14)  ' N20 - Protective Payment Code
    qcData(qcwr, 23) = srcCache(24, 40)  ' AN24 - TANF Days
End Sub

Private Sub ExtractPLInfo()
    Dim j As Long, ln As Long
    ln = 0
    
    For j = 30 To 44 Step 2
        If IsEmpty(srcCache(j, 1)) Or srcCache(j, 1) = "" Then Exit For
        
        plwr = plwr + 1
        ln = ln + 1
        
        plData(plwr, 1) = revidval
        plData(plwr, 2) = srcCache(j, 1)      ' A - Person number
        plData(plwr, 3) = srcCache(j, 4)      ' D - Case_Afl_First_Code
        plData(plwr, 4) = srcCache(j, 5)      ' E - Case_Afl_Second_Code
        plData(plwr, 5) = srcCache(j, 7)      ' G - Deprivation_Code
        plData(plwr, 6) = srcCache(j, 10)     ' J - Relationship_Payment_Code
        plData(plwr, 7) = Val(srcCache(j, 13)) ' M - Age
        plData(plwr, 8) = srcCache(j, 16)     ' P - Gender
        plData(plwr, 9) = srcCache(j, 18)     ' R - Race
        plData(plwr, 10) = srcCache(j, 20)    ' T - Citizenship
        plData(plwr, 11) = srcCache(j, 22)    ' V - Education Level
        plData(plwr, 12) = srcCache(j, 25)    ' Y - Reset Code
        plData(plwr, 13) = srcCache(j, 29)    ' AC - Work Activity Code
        plData(plwr, 14) = srcCache(j, 32)    ' AF - Referral Days
        plData(plwr, 15) = srcCache(j, 37)    ' AK - Marital Status Code
        plData(plwr, 16) = srcCache(j, 40)    ' AN - Program Status Code
        plData(plwr, 17) = srcCache(j, 43)    ' AQ - TE Code
        plData(plwr, 18) = ln                  ' Line No
    Next j
End Sub

Private Sub ExtractHHInc()
    Dim j As Long, k As Long
    
    For j = 50 To 56 Step 2
        If IsEmpty(srcCache(j, 3)) Or srcCache(j, 3) = "" Then Exit For
        
        For k = 7 To 37 Step 10
            If IsEmpty(srcCache(j, k)) Or srcCache(j, k) = "" Then Exit For
            
            hiwr = hiwr + 1
            hiData(hiwr, 1) = revidval
            hiData(hiwr, 2) = srcCache(j, k + 4)  ' Amount of income
            hiData(hiwr, 4) = srcCache(j, 3)       ' C - Person number
            hiData(hiwr, 5) = srcCache(j, k)       ' Type of income
        Next k
    Next j
End Sub

Private Sub ExtractErrFind()
    Dim j As Long, ln As Long
    ln = 0
    
    For j = 61 To 67 Step 2
        If IsEmpty(srcCache(j, 6)) Or srcCache(j, 6) = "" Then Exit For
        
        efwr = efwr + 1
        ln = ln + 1
        
        efData(efwr, 1) = revidval
        efData(efwr, 2) = srcCache(j, 6)   ' F - Error_Findings Code
        efData(efwr, 3) = srcCache(j, 44)  ' AR - Occurrence_Period_Code
        efData(efwr, 4) = srcCache(j, 30)  ' AD - Discovery_Code
        efData(efwr, 5) = srcCache(j, 34)  ' AH - Verification_Code
        efData(efwr, 6) = srcCache(j, 20)  ' T - Client_Agency_Code
        efData(efwr, 7) = srcCache(j, 3)   ' C - Optional
        efData(efwr, 8) = srcCache(j, 24)  ' X - Dollar_Amount
        
        ' Occurrence Date (AL = col 38)
        If IsDate(srcCache(j, 38)) Then
            Dim occDate As Date: occDate = srcCache(j, 38)
            efData(efwr, 9) = DateSerial(Year(occDate), Month(occDate), 1)
        End If
        
        efData(efwr, 10) = srcCache(j, 15)  ' O - Nature_Code
        efData(efwr, 11) = srcCache(j, 10)  ' J - Element_Code
        efData(efwr, 12) = ln                ' Line number
    Next j
End Sub

' ============================================================================
' HELPER FUNCTIONS
' ============================================================================

Private Function CleanValue(v As Variant, defaultVal As String) As String
    ' Returns cleaned string or default if empty/invalid
    If IsError(v) Then
        CleanValue = defaultVal
    ElseIf IsEmpty(v) Then
        CleanValue = defaultVal
    ElseIf Len(Trim(CStr(v))) = 0 Then
        CleanValue = defaultVal
    ElseIf InStr(CStr(v), "-") > 0 Then
        CleanValue = defaultVal
    Else
        CleanValue = CStr(v)
    End If
End Function

Private Function TrimArray(arr As Variant, rowCount As Long, colCount As Long) As Variant
    ' Returns a properly sized copy of the array (removes unused pre-allocated rows)
    Dim result() As Variant
    Dim r As Long, c As Long
    
    If rowCount = 0 Then
        TrimArray = Array()
        Exit Function
    End If
    
    ReDim result(1 To rowCount, 1 To colCount)
    For r = 1 To rowCount
        For c = 1 To colCount
            result(r, c) = arr(r, c)
        Next c
    Next r
    
    TrimArray = result
End Function

Private Sub TransferTableToAccess(cnt As ADODB.Connection, ws As Worksheet, _
                                   tableName As String, maxRows As Long)
    ' Transfers worksheet data to Access table using optimized recordset operations
    If maxRows = 0 Then Exit Sub
    
    Dim rs As ADODB.Recordset
    Dim headerArr As Variant, dataArr As Variant
    Dim FieldMap() As Long
    Dim lastCol As Long, r As Long, c As Long
    Dim hName As String, fldIdx As Long
    
    ' Find last column with header
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Read all data at once
    headerArr = ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol)).Value
    dataArr = ws.Range(ws.Cells(2, 1), ws.Cells(maxRows + 1, lastCol)).Value
    
    ' Open recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open tableName, cnt, adOpenKeyset, adLockOptimistic, adCmdTableDirect
    
    ' Pre-calculate column-to-field mapping
    ReDim FieldMap(1 To lastCol)
    For c = 1 To lastCol
        FieldMap(c) = -1  ' Default: no match
        hName = Trim(CStr(headerArr(1, c)))
        
        If hName <> "" Then
            ' Normalize header (replace spaces with underscores, lowercase)
            Dim normH As String
            normH = Replace(LCase(hName), " ", "_")
            
            For fldIdx = 0 To rs.Fields.Count - 1
                If LCase(rs.Fields(fldIdx).Name) = normH Or _
                   LCase(rs.Fields(fldIdx).Name) = LCase(hName) Then
                    FieldMap(c) = fldIdx
                    Exit For
                End If
            Next fldIdx
        End If
    Next c
    
    ' Batch insert rows - NO silent error suppression
    Dim vValue As Variant
    For r = 1 To UBound(dataArr, 1)
        rs.AddNew
        For c = 1 To UBound(dataArr, 2)
            If FieldMap(c) >= 0 Then
                vValue = dataArr(r, c)
                If Not IsError(vValue) And Not IsEmpty(vValue) Then
                    If Len(CStr(vValue)) > 0 Then
                        ' Let errors propagate - will trigger DBError handler
                        ' which rolls back transaction and reports the issue
                        rs.Fields(FieldMap(c)).Value = vValue
                    End If
                End If
            End If
        Next c
        rs.Update
    Next r
    
    rs.Close
    Set rs = Nothing
End Sub

