' ============================================================================
' MA Negative Edit Check Optimized
' ============================================================================
' This macro does the following:
'   - Takes a list of review numbers (and their sample months + examiner numbers) from the
'     currently active "control" worksheet.
'   - Locates each examiner's review workbook on the network drive using a predictable
'     folder structure.
'   - Opens each review workbook READ-ONLY, pulls a fixed set of values from a known
'     template layout, and stores them into 1 output table (CaseReview_dtl).
'   - Writes that table into a temporary Excel file, then transfers it into a
'     new Access database (.mdb) in ONE transaction (all-or-nothing).
'
' The main speed trick is to read big blocks *once* into arrays, work in memory, then write big blocks *once* at the end.
'
' Key points:
'   - Batched Input/Output: `Range(...).Value` into Variant arrays is thousands of times faster than reading/writing cells one at a time.
'   - Template layout indexing: Values like `srcCache(15, 12)` mean "Row 15, Column 12" of the cached template range.
'   - MA Negative only has one output table (CaseReview_dtl) with 19 columns.
' ============================================================================
' REQUIREMENTS:
'   - Reference: Microsoft ActiveX Data Objects 6.1 Library (or 2.8+)
'       If you get an error like "User Defined Type not defined" on ADODB.Connection or the macro fails on opening the database with a provider error,
'       Just check the "Microsoft ActiveX Data Objects" Library box in VBA Editor > Tools > References in the options.
'   - Access Database Engine installed for ACE OLEDB 12.0 provider
'       Same here, if the macro errors with something like "Provider cannot be found" / "not registered",
'       double check that Access is installed properly.
'   If the button works, don't worry about this.
' ============================================================================
Option Explicit

' Output row counter
Private drwr As Long

' Pre-allocated output array (sized for max expected records)
Private drData() As Variant   ' CaseReview_dtl

' Column count for the table
Private Const DR_COLS As Long = 19   ' Columns A-S in CaseReview_dtl

' ==========================================================================
' SOURCE TEMPLATE LAYOUT CONSTANTS
' ==========================================================================
' We read a single rectangle (`SRC_CACHE_RANGE`) into memory, then use row/column indexes into that array.
' If the template changes, these references must be updated. Check here first if the macro breaks after an update.

' Source data range - covers all used rows (highest data row is 56)
' Rows used: 11, 15, 19, 25, 40, 56
Private Const SRC_CACHE_RANGE As String = "A1:AF60"

' Source workbook data cache
Private srcCache As Variant

' Headers for the output table
Private Const HEADERS As String = "ReviewNo,SampleMonth,ReviewerNo,StateCode,CountyCode,CaseNo," & _
    "CaseCategoryCode,ProgramStatus,GrantGroup,AgencyDecisionDate,AgencyActionDate," & _
    "ReviewCategory,ActionTypeCode,HearingReqCode,ReasonForActionCode," & _
    "EligibilityRequirementCode,FieldInvestigationCode,DispositionCode,PostReviewStatusCode"

Sub Find_Write_Database_Files()
    On Error GoTo ErrorHandler
    
    ' ========================================================================
    ' 1. PERFORMANCE OPTIMIZATION - Disable all Excel overhead
    ' ========================================================================
    With Application
        .ScreenUpdating = False     ' avoid repainting the Excel UI after each operation
        .DisplayStatusBar = True    ' we still show a status bar to the user for feedback.
        .EnableEvents = False       ' prevent workbook/worksheet event macros from firing mid-run
        .Calculation = xlCalculationManual  ' prevent formula recalculation after each write (we do very few writes, but this is still a safe default for most runs)
        ' we restore these settings on exit (see CleanExit) so Excel doesn't "feel broken" to the next person using it.
    End With
    
    ' Variable declarations
    Dim i As Long, j As Long
    Dim maxrow As Long, maxrowex As Long
    Dim program As String, exname As String, exnumstr As String
    Dim reviewtxt As String
    Dim pathdir As String, BasePath As String, FinalFilePath As String, FileNameFound As String
    Dim monthstr As String, mName As String, yStr As String
    Dim sPath As String, SrceFile As String, exceloutfile As String, databasename As String
    Dim DLetter As String
    Dim filenum As Integer
    Dim processedCount As Long
    
    ' Workbook/Worksheet objects
    Dim thissht As Worksheet
    Dim inWB As Workbook, outWB As Workbook
    Dim inWS As Worksheet
    Dim outWS As Worksheet
    
    ' Network objects
    Dim WshNetwork As Object, oDrives As Object
    
    ' ADO objects
    Dim cnt As ADODB.Connection
    
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
        Err.Raise 9999, "Find_Write_Database_Files", _
            "Network Drive to Examiner Files NOT found." & vbCrLf & "Contact Valerie or Wes."
    End If
    
    pathdir = DLetter & "Schedules by Examiner Number\"
    If Dir(pathdir, vbDirectory) = "" Then
        Err.Raise 9999, "Find_Write_Database_Files", _
            "Examiner Directory not found: " & pathdir
    End If
    
    ' ========================================================================
    ' 3. LOAD ALL INPUT DATA INTO MEMORY (Single read operations)
    ' ========================================================================
    ' Accessing `Cells(i, j)` in a loop is slow because every call goes through Excel's COM interface. Reading the whole block once is fast, then indexing the array is fast.
    ' Another small-but-important improvement vs the original: we use End(xlUp) from the bottom of the sheet.
    ' End(xlDown) can stop early if there are blanks, which makes the macro "randomly" skip cases.
    maxrow = thissht.Range("E" & thissht.Rows.Count).End(xlUp).Row
    maxrowex = thissht.Range("L" & thissht.Rows.Count).End(xlUp).Row
    
    ' Load review data: E=ReviewNum, F=Month, G=ExaminerNum
    vInput = thissht.Range("E1:G" & maxrow).Value
    ' Load examiner lookup: K=Name, L=Number
    vExLookup = thissht.Range("K1:L" & maxrowex).Value
    
    ' ========================================================================
    ' 4. PRE-ALLOCATE OUTPUT ARRAY (Avoids incremental resizing)
    ' ========================================================================
    ' Growing arrays inside loops (ReDim Preserve) is extremely expensive. 
    ' Instead, we allocate "big enough" arrays once, fill them, then trim to the actual used size right before writing them to the output sheets (see TrimArray()).
    ' MA Negative only has one output table (CaseReview_dtl), so we only need one array.
    Dim maxCases As Long
    maxCases = maxrow - 1  ' Estimate max records
    
    ReDim drData(1 To maxCases, 1 To DR_COLS)
    
    ' Initialize counter (arrays are 1-based, row 1 = first data row)
    drwr = 0
    processedCount = 0
    
    ' ========================================================================
    ' 5. MAIN PROCESSING LOOP
    ' ========================================================================
    ' General flow per row:
    '   Read review + examiner info from the input arrays
    '   Build the expected folder path for that examiner/program/month
    '   Find the specific review workbook with `Dir()` (MA Negative uses a different file naming pattern than TANF)
    '   Open workbook read-only, cache its data range into `srcCache`
    '   Extract data into the output array
    '   Close workbook and continue
    For i = 2 To maxrow
        ' Update status every record
        Application.StatusBar = "Processing " & (i - 1) & "/" & (maxrow - 1) & _
            " (" & Format((i - 1) / (maxrow - 1), "0%") & ")"
        DoEvents
        
        ' ---------------------------------------------------------------------
        ' A. Parse and validate review number
        ' ---------------------------------------------------------------------
        ' Strip the leading zero from the review number
        reviewtxt = Trim(CStr(vInput(i, 1)))
        If Left(reviewtxt, 1) = "0" Then reviewtxt = Mid(reviewtxt, 2)
        
        ' Only process MA Negative reviews (start with "8")
        If Left(reviewtxt, 1) <> "8" Then GoTo NextIteration
        program = "MA Negative"
        
        ' ---------------------------------------------------------------------
        ' B. Find examiner name (array lookup - no Excel calls)
        ' ---------------------------------------------------------------------
        ' A simple in-memory scan of the examiner array is fast and predictable. VLOOKUP would add more Excel calls (slow) and can be brittle if the range moves or changes.
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
        ' The folder naming convention uses "Name - N" where N is 1-2 digits without a leading 0. Input may contain "01" due to excel auto formatting but folder is "... - 1".
        exnumstr = Format(currentExNum, "00")
        If Left(exnumstr, 1) = "0" Then exnumstr = Right(exnumstr, 1)
        
        ' ---------------------------------------------------------------------
        ' C. Parse sample month
        ' ---------------------------------------------------------------------
        ' Expected: YYYYMM (6 digits). We use this to build "Review Month {MonthName} {YYYY}" which matches the network folder structure.
        monthstr = Trim(CStr(vInput(i, 2)))
        If Len(monthstr) < 6 Then GoTo NextIteration
        
        yStr = Left(monthstr, 4)
        mName = MonthName(Val(Right(monthstr, 2)))
        
        ' ---------------------------------------------------------------------
        ' D. Build path and find workbook
        ' ---------------------------------------------------------------------
        ' We use the Dir() function to find the case folder and workbook. It is built into VBA, fast, and avoids extra dependencies, like the custom class module used in the original version.
        ' NOTE: This assumes the network folder structure is correct and the case folder exists. If the folder structure changes, this will need to be updated.
        '
        ' Folder structure: {program}\[FFY archive]\Review Month {MonthName} {YYYY}\{case folder}\review.xlsm
        ' MA reviews are long-running, so the monthly folder may have been archived to "FFY 20XX - [exnumstr]"
        ' before we run the edit check. We try both active and archived locations.
        
        Dim ffyYear As Long
        Dim mNum As Long
        Dim CaseFolderName As String
        Dim CaseFolderPath As String
        
        mNum = Val(Right(monthstr, 2))
        ' FFY runs Oct-Sep: Oct-Dec = next year's FFY, Jan-Sep = current year's FFY
        If mNum >= 10 Then
            ffyYear = Val(yStr) + 1
        Else
            ffyYear = Val(yStr)
        End If
        
        FileNameFound = ""
        CaseFolderName = ""
        
        ' Try active location first (no FFY archive folder)
        BasePath = pathdir & exname & " - " & exnumstr & "\" & program & "\" & _
                   "Review Month " & mName & " " & yStr & "\"
        
        If Dir(BasePath, vbDirectory) <> "" Then
            CaseFolderName = Dir(BasePath & reviewtxt & " - *", vbDirectory)
        End If
        
        ' If not found, try FFY archive folder
        If CaseFolderName = "" Then
            BasePath = pathdir & exname & " - " & exnumstr & "\" & program & "\" & _
                       "FFY " & ffyYear & " - " & exnumstr & "\" & _
                       "Review Month " & mName & " " & yStr & "\"
            If Dir(BasePath, vbDirectory) <> "" Then
                CaseFolderName = Dir(BasePath & reviewtxt & " - *", vbDirectory)
            End If
        End If
        
        If CaseFolderName = "" Then GoTo NextIteration
        If (GetAttr(BasePath & CaseFolderName) And vbDirectory) <> vbDirectory Then GoTo NextIteration
        
        CaseFolderPath = BasePath & CaseFolderName & "\"
        FileNameFound = Dir(CaseFolderPath & "Review Number " & reviewtxt & " Month " & monthstr & " Examiner*.xls*")
        
        If FileNameFound = "" Then GoTo NextIteration
        
        FinalFilePath = CaseFolderPath & FileNameFound
        
        ' ---------------------------------------------------------------------
        ' E. Open and process workbook
        ' ---------------------------------------------------------------------
        'Review workbooks contain external links or prompts. We want a non-interactive, predictable run. This also prevents accidental edits.
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
        ' F. BATCH READ: Load all source data into cache array
        ' ---------------------------------------------------------------------
        '   - This is the ONE Excel read we do per case. Everything else reads from `srcCache`.
        '   - `SRC_CACHE_RANGE` covers every cell we reference in the rest of the extraction logic.
        '   - Reminder: `srcCache(r, c)` is 1-based and uses A=1, B=2, ... AF=32 for columns.
        srcCache = inWS.Range(SRC_CACHE_RANGE).Value
        
        ' ---------------------------------------------------------------------
        ' G. Extract data to output array
        ' ---------------------------------------------------------------------
        ' The extraction logic is in a subroutine to keep the main loop readable and to centralize the "template cell mapping" logic.
        ' MA Negative only extracts one table (CaseReview_dtl), unlike TANF which extracts 5 tables.
        Call ExtractCaseReview
        
        inWB.Close False
        Set inWB = Nothing
        processedCount = processedCount + 1
        
NextIteration:
    Next i
    
    ' ========================================================================
    ' 6. WRITE OUTPUT ARRAY TO EXCEL
    ' ========================================================================
    ' Excel writes are expensive. Writing one big block at the end is much faster than writing individual cells as we process each case.
    ' We write starting at row 2 because row 1 contains the headers.
    ' NOTE: We call TrimArray(...) so we only write the populated rows (and not all the extra pre-allocated blank rows).
    Application.StatusBar = "Writing to Excel..."
    DoEvents
    
    ' Create new workbook with headers
    Set outWB = Workbooks.Add(1)
    Set outWS = outWB.ActiveSheet
    exceloutfile = sPath & "\MA Negative Database Input " & Format(Date, "mm-dd-yyyy") & ".xlsx"
    
    ' Write headers
    Dim headerArr() As String
    headerArr = Split(HEADERS, ",")
    Dim h As Long
    For h = 0 To UBound(headerArr)
        outWS.Cells(1, h + 1).Value = headerArr(h)
    Next h
    
    ' Write data array (single range write per table)
    If drwr > 0 Then
        outWS.Range("A2").Resize(drwr, DR_COLS).Value = TrimArray(drData, drwr, DR_COLS)
    End If
    
    ' Save workbook
    If Dir(exceloutfile) <> "" Then Kill exceloutfile
    outWB.SaveAs Filename:=exceloutfile, FileFormat:=51
    
    ' ========================================================================
    ' 7. DATABASE TRANSFER WITH TRANSACTION
    ' ========================================================================
    ' We use a transaction to write the data to the Access database. This is more efficient than writing one row at a time.
    ' If any row fails, we roll back everything so the output database is not partially populated, saving us from deleting the partial database when starting over.
    Application.StatusBar = "Transferring to Access database..."
    DoEvents
    
    ' Prepare database file
    ' NOTE: MA Negative uses a different source path for the blank database (on the network drive, not in the local FO Databases folder)
    SrceFile = DLetter & "HQ - Data Entry\Create FO Databases\FO Databases\MA_Neg_Blank.mdb"
    databasename = sPath & "\MA NEG1 " & Format(Date, "mm-dd-yyyy") & ".mdb"
    filenum = 1
    Do Until Dir(databasename) = ""
        filenum = filenum + 1
        databasename = sPath & "\MA NEG" & filenum & " " & Format(Date, "mmddyyyy") & ".mdb"
    Loop
    FileCopy SrceFile, databasename
    
    ' Open connection with transaction
    Set cnt = New ADODB.Connection
    cnt.Open "Provider=Microsoft.Ace.OLEDB.12.0;Data Source=" & databasename & ";"
    cnt.BeginTrans  ' One big transaction = faster + all-or-nothing if something goes wrong
    
    On Error GoTo DBError
    
    ' Transfer table
    Call TransferTableToAccess(cnt, outWS, "CaseReview_dtl", drwr)
    
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
    
    ' Clear array from memory
    Erase drData
    
    MsgBox "Processing Complete!" & vbCrLf & vbCrLf & _
           "Cases Processed: " & processedCount & vbCrLf & _
           "Database: " & databasename, vbInformation, "MA Negative Edit Check V2"

CleanExit:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

DBError:
    ' Rollback ALL changes - no partial data
    ' This is intentional: partial databases unnecessary given we get the intermediate excel file if there is an error.
    cnt.RollbackTrans
    MsgBox "DATABASE WRITE FAILED - NO DATA SAVED" & vbCrLf & vbCrLf & _
           "Error: " & Err.Description & vbCrLf & _
           "Error #: " & Err.Number & vbCrLf & vbCrLf & _
           "The transaction has been rolled back." & vbCrLf & _
           "Please review the source data and retry.", vbCritical
    Resume CleanExit

ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & _
           "Source: " & Err.Source, vbCritical, "MA Negative Edit Check V2 - Error"
    If Not inWB Is Nothing Then inWB.Close False
    If Not outWB Is Nothing Then outWB.Close False
    If Not cnt Is Nothing Then
        If cnt.State = adStateOpen Then cnt.Close
    End If
    Resume CleanExit
End Sub

' ============================================================================
' DATA EXTRACTION SUBROUTINE
' ============================================================================
' Uses srcCache array - no Excel calls

Private Sub ExtractCaseReview()
' PURPOSE:
'   - Writes one row to the CaseReview_dtl table for the current case.
'   - All fields come from `srcCache` using template cell mappings. If a column moves on the template, adjust the `srcCache(row, col)` indexes here.
    drwr = drwr + 1
    
    drData(drwr, 1) = srcCache(15, 12)                                      ' A - ReviewNo (L15)
    
    ' Sample Month (T11 = col 20) - parse MMYYYY format to date
    ' The template stores sample month as numeric text; we normalize to a first day of month because Access date fields expect dates, not strings.
    Dim sm As Variant: sm = srcCache(11, 20)
    If IsNumeric(sm) And Len(CStr(sm)) >= 6 Then
        Dim smStr As String: smStr = Format(sm, "000000")
        drData(drwr, 2) = DateSerial(Val(Right(smStr, 4)), Val(Left(smStr, 2)), 1)  ' B - SampleMonth
    End If
    
    drData(drwr, 3) = CStr(srcCache(11, 28)) & CStr(srcCache(11, 29))       ' C - ReviewerNo (AB11 & AC11)
    drData(drwr, 4) = srcCache(15, 2)                                       ' D - StateCode (B15)
    drData(drwr, 5) = srcCache(15, 7)                                       ' E - CountyCode (G15)
    
    ' Case Number (S15 = col 19) - remove leading "A" if present
    ' Some case numbers have a leading "A" prefix that should be stripped for database consistency.
    Dim caseNum As Variant: caseNum = srcCache(15, 19)
    If Left(CStr(caseNum), 1) = "A" Then
        drData(drwr, 6) = Mid(CStr(caseNum), 2)                             ' F - CaseNo
    Else
        drData(drwr, 6) = caseNum                                           ' F - CaseNo
    End If
    
    drData(drwr, 7) = Trim(CStr(srcCache(15, 28)))                          ' G - CaseCategoryCode (AB15)
    drData(drwr, 8) = srcCache(19, 2)                                       ' H - ProgramStatus (B19)
    drData(drwr, 9) = srcCache(19, 6)                                       ' I - GrantGroup (F19)
    
    ' Agency Decision Date (J19 = col 10) - only if valid date
    ' Date fields are optional - only write if the cell contains a valid date value.
    If IsDate(srcCache(19, 10)) Then
        drData(drwr, 10) = srcCache(19, 10)                                 ' J - AgencyDecisionDate
    End If
    
    ' Agency Action Date (S19 = col 19) - only if valid date
    ' Date fields are optional - only write if the cell contains a valid date value.
    If IsDate(srcCache(19, 19)) Then
        drData(drwr, 11) = srcCache(19, 19)                                 ' K - AgencyActionDate
    End If
    
    drData(drwr, 12) = srcCache(19, 28)                                     ' L - ReviewCategory (AB19)
    drData(drwr, 13) = srcCache(19, 32)                                     ' M - ActionTypeCode (AF19)
    drData(drwr, 14) = CleanValue(srcCache(25, 3), "B")                     ' N - HearingReqCode (C25)
    drData(drwr, 15) = CleanValue(srcCache(25, 17), "BB")                   ' O - ReasonForActionCode (Q25)
    drData(drwr, 16) = CleanValue(srcCache(40, 3), "B")                     ' P - EligibilityRequirementCode (C40)
    drData(drwr, 17) = CleanValue(srcCache(56, 3), "B")                     ' Q - FieldInvestigationCode (C56)
    drData(drwr, 18) = CleanValue(srcCache(56, 13), "B")                    ' R - DispositionCode (M56)
    drData(drwr, 19) = CleanValue(srcCache(56, 25), "B")                    ' S - PostReviewStatusCode (Y56)
End Sub

' ============================================================================
' HELPER FUNCTIONS
' ============================================================================

Private Function CleanValue(v As Variant, defaultVal As String) As String
' PURPOSE:
'   - The template sometimes contains blanks, errors (#N/A), or placeholders like "-".
' RETURNS:
'   - A safe string: either the cleaned value, or the provided default.
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
' PURPOSE:
'   - We pre-allocate arrays larger than we need for speed. This function copies just the populated portion into a new array so the Excel `.Resize(...).Value = ...` write doesn't include any blank rows.
'   - `ReDim Preserve` on a multi-dimensional array is slow; copying is clear and fast enough given we only do it once per table at the end.
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
' PURPOSE:
'   - Bulk-loads one worksheet table into the matching Access table. This is the biggest change to the database logic from the original version.
'   - Using Excel as an ODBC/ISAM data source is slow and error-prone (as we've seen). Recordset inserts keep the data flow in memory and allow us to use a transaction.
'   - This sub assumes row 1 contains headers that correspond to Access field names. We build a column→field mapping once, then loop the data rows and set values by index.
'   - We intentionally do NOT suppress errors during writes. If a value doesn't fit the database schema, we want to stop and rollback (so we don't hide a half-bad database) and report the error.
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
    rs.CursorLocation = adUseClient ' client-side cursor is usually faster for batch inserts like this
    rs.Open tableName, cnt, adOpenKeyset, adLockOptimistic, adCmdTableDirect
    ' adCmdTableDirect tells ADO we're opening a table, not running a SQL query (less overhead).
    
    ' Pre-calculate column-to-field mapping
    ' FieldMap(c) stores the Access field index that matches worksheet column c.
    ' If no match is found, FieldMap(c) stays -1 and we skip that worksheet column.
    ReDim FieldMap(1 To lastCol)
    For c = 1 To lastCol
        FieldMap(c) = -1  ' Default: no match
        hName = Trim(CStr(headerArr(1, c)))
        
        If hName <> "" Then
            ' Normalize header (replace spaces with underscores, lowercase)
            ' Excel headers may contain spaces ("Run Date") but Access field names often use underscores ("run_date"). We can accept either by normalizing and comparing both forms.
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
    
    ' Batch insert rows
    ' Excel might contain empty variants or error values (ex: #VALUE!). Writing those directly can raise ADO type conversion errors. We treat them as "no value" and let Access defaults / null handling apply.
    Dim vValue As Variant
    For r = 1 To UBound(dataArr, 1)
        rs.AddNew
        For c = 1 To UBound(dataArr, 2)
            If FieldMap(c) >= 0 Then
                vValue = dataArr(r, c)
                If Not IsError(vValue) And Not IsEmpty(vValue) Then
                    If Len(CStr(vValue)) > 0 Then
                        ' Let errors propagate; this will trigger the DBError handler,
                        ' which rolls back the transaction and reports the issue.
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

