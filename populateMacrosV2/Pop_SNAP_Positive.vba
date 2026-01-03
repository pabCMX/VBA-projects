Attribute VB_Name = "Pop_SNAP_Positive"
' ============================================================================
' Pop_SNAP_Positive - SNAP Positive Schedule Population Module
' ============================================================================
' WHAT THIS MODULE DOES:
'   This module reads data from BIS (Benefits Information System) delimited
'   files and populates SNAP Positive (food assistance approval) review
'   schedule templates. It's the core data extraction and mapping logic
'   for SNAP Positive reviews.
'
' HOW THE POPULATION PROCESS WORKS:
'   1. The main module (Pop_Main) opens the BIS file and schedule template
'   2. This module finds the review number in the BIS "Case" sheet
'   3. Finds all household members in the BIS "Individual" sheet
'   4. Sorts members (head of household first, then by line number)
'   5. Maps BIS fields to schedule cells (FS Workbook and Schedule)
'   6. Populates checkbox values based on eligibility codes
'
' BIS FILE STRUCTURE:
'   The BIS delimited file has two worksheets:
'   - "Case": One row per case with case-level data (address, dates, totals)
'   - "Individual": One row per household member with person-level data
'
' KEY CONCEPTS:
'   - Line Number (LN): BIS person identifier within household (01, 02, etc.)
'   - Eligibility Status: EM=Eligible Member, NM=Non-Member, EB=ABAWD, etc.
'   - ETP Code: Employment & Training Program status code
'   - ABAWD: Able-Bodied Adult Without Dependents (work requirement)
'   - SUA: Standard Utility Allowance
'   - LIHEAP: Low Income Home Energy Assistance Program
'
' WORKBOOK REFERENCES:
'   - wb_bis: The BIS delimited file (opened by Pop_Main)
'   - wb: The output schedule workbook being populated
'   - "FS Workbook": The SNAP workbook sheet with computation elements
'   - review_number sheet: The main SNAP schedule (named by review number)
'
' ELEMENT NUMBERS:
'   The code references "Element XXX" which are federal QC schedule items:
'   - Element 110: Birthdate verification
'   - Element 111: Student status (age 18+ in school)
'   - Element 130: Citizenship status
'   - Element 150: Household composition
'   - Element 160: E&T (Employment & Training) participation
'   - Element 161: ABAWD status
'   - Element 331-343: Unearned income types (RSDI, VA, SSI, UC, etc.)
'
' DEPENDENCIES:
'   - Common_Utils: For helper functions (line number formatting)
'   - Config_Settings: For income frequency multipliers
'   - Pop_Main: Sets up wb_bis and wb workbook references
'
' GLOBAL VARIABLES (from Pop_Main):
'   - wb_bis: BIS workbook
'   - wb: Output schedule workbook
'   - program: Current program type string
'
' CHANGE LOG:
'   2026-01-03  Created from Populate_snap_pos_delimited_mod.vba
'               Added Option Explicit, V2 comments, error handling
' ============================================================================

Option Explicit

' ============================================================================
' MAIN POPULATION FUNCTION
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: PopulateSNAPPositiveDelimited
' ----------------------------------------------------------------------------
' PURPOSE:
'   Main entry point for populating a SNAP Positive schedule from BIS data.
'   Called by Pop_Main after the BIS file and schedule template are opened.
'
' PRE-CONDITIONS:
'   - wb_bis (Public in Pop_Main) contains the opened BIS delimited file
'   - wb (Public in Pop_Main) contains the schedule workbook to populate
'   - program (Public in Pop_Main) is set to "FS Positive" or "FS Supplemental"
'
' POST-CONDITIONS:
'   - Schedule sheets are populated with BIS data
'   - Checkboxes are set based on eligibility conditions
'   - Computation sections pre-filled where possible
'
' CALLED FROM:
'   Pop_Main.PopulateSNAPPositive()
' ----------------------------------------------------------------------------
Public Sub PopulateSNAPPositiveDelimited()
    On Error GoTo ErrorHandler
    
    ' ========================================================================
    ' VARIABLE DECLARATIONS
    ' ========================================================================
    
    ' Workbook and worksheet references
    Dim wbSch As Workbook           ' Schedule workbook being populated
    Dim wsBISCase As Worksheet      ' BIS Case worksheet
    Dim wsBISInd As Worksheet       ' BIS Individual worksheet
    Dim wsFSWorkbook As Worksheet   ' FS Workbook sheet in schedule
    Dim wsSchedule As Worksheet     ' Main schedule sheet (named by review #)
    
    ' Data location variables
    Dim reviewNumber As String      ' The review number (e.g., "5012345")
    Dim caseRow As Range            ' Row in BIS Case sheet for this review
    Dim maxrowBISCase As Long       ' Last row in BIS Case sheet
    Dim maxrowBISInd As Long        ' Last row in BIS Individual sheet
    Dim lastColBISCase As Long      ' Last column in BIS Case sheet
    Dim lastColBISInd As Long       ' Last column in BIS Individual sheet
    Dim bisIndStartRow As Long      ' First row of this case in Individual sheet
    Dim bisIndStopRow As Long       ' Last row of this case in Individual sheet
    
    ' Income frequency multiplier array (for converting to monthly)
    Dim incomeFreq As Variant
    incomeFreq = GetIncomeMultiplierArray()
    
    ' Processing variables
    Dim ws As Worksheet
    Dim i As Long, j As Long
    Dim temp As Variant
    Dim writeRow As Long
    
    ' ========================================================================
    ' 1. SETUP - GET WORKBOOK REFERENCES
    ' ========================================================================
    
    ' Get reference to schedule workbook (from Pop_Main global)
    Set wbSch = wb
    
    ' Find the review number by looking for a sheet with a numeric name > 1000
    ' (Schedule sheets are named with the review number)
    For Each ws In wbSch.Worksheets
        If Val(ws.Name) > 1000 Then
            reviewNumber = ws.Name
            Exit For
        End If
    Next ws
    
    If reviewNumber = "" Then
        MsgBox "Could not find schedule sheet (review number > 1000)", vbCritical
        Exit Sub
    End If
    
    ' Set up worksheet references
    Set wsBISCase = wb_bis.Worksheets("Case")
    Set wsBISInd = wb_bis.Worksheets("Individual")
    Set wsFSWorkbook = wbSch.Worksheets("FS Workbook")
    Set wsSchedule = wbSch.Worksheets(reviewNumber)
    
    ' ========================================================================
    ' 2. FIND CASE IN BIS FILE
    ' ========================================================================
    
    ' Find extents of BIS Case worksheet
    maxrowBISCase = wsBISCase.Cells(wsBISCase.Rows.Count, 1).End(xlUp).Row
    lastColBISCase = wsBISCase.Cells(1, wsBISCase.Columns.Count).End(xlToLeft).Column
    
    ' Find the row for this review number in BIS Case sheet
    ' Review number is in column C
    With wsBISCase.Range("C2:C" & maxrowBISCase)
        Set caseRow = .Find(Val(reviewNumber), LookIn:=xlValues)
    End With
    
    If caseRow Is Nothing Then
        MsgBox "Review number " & reviewNumber & " not found in BIS Case file", vbExclamation
        Exit Sub
    End If
    
    ' ========================================================================
    ' 3. FIND INDIVIDUALS FOR THIS CASE
    ' ========================================================================
    
    ' Find extents of BIS Individual worksheet
    maxrowBISInd = wsBISInd.Cells(wsBISInd.Rows.Count, 1).End(xlUp).Row
    lastColBISInd = wsBISInd.Cells(1, wsBISInd.Columns.Count).End(xlToLeft).Column
    
    ' Find the start and stop rows for this review number in Individual sheet
    ' Review number is in column C of Individual sheet
    bisIndStartRow = 0
    bisIndStopRow = 0
    
    For i = 2 To maxrowBISInd
        If Val(wsBISInd.Range("C" & i).Value) = Val(reviewNumber) Then
            If bisIndStartRow = 0 Then
                bisIndStartRow = i
            End If
            bisIndStopRow = i
        ElseIf bisIndStartRow > 0 Then
            ' We've passed all rows for this review number
            Exit For
        End If
    Next i
    
    If bisIndStartRow = 0 Then
        ' No individuals found - unusual but may be valid for some cases
        ' Continue with case-level data only
    End If
    
    ' ========================================================================
    ' 4. SORT INDIVIDUALS (HOH FIRST, THEN BY LINE NUMBER)
    ' ========================================================================
    
    If bisIndStopRow > bisIndStartRow Then
        Call SortIndividuals(wsBISInd, bisIndStartRow, bisIndStopRow, lastColBISInd)
    End If
    
    ' ========================================================================
    ' 5. POPULATE CASE-LEVEL DATA
    ' ========================================================================
    
    Call PopulateCaseLevelData(wsFSWorkbook, wsSchedule, wsBISCase, caseRow.Row, incomeFreq)
    
    ' ========================================================================
    ' 6. POPULATE INDIVIDUAL-LEVEL DATA
    ' ========================================================================
    
    If bisIndStartRow > 0 Then
        Call PopulateIndividualData(wsFSWorkbook, wsSchedule, wsBISInd, _
                                    bisIndStartRow, bisIndStopRow, reviewNumber, incomeFreq)
    End If
    
    ' ========================================================================
    ' 7. POPULATE SCHEDULE SECTION 4 (HOUSEHOLD CHARACTERISTICS)
    ' ========================================================================
    
    If bisIndStartRow > 0 Then
        Call PopulateScheduleSection4(wsSchedule, wsBISInd, bisIndStartRow, bisIndStopRow)
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in PopulateSNAPPositiveDelimited: " & Err.Description, vbCritical
    LogError "PopulateSNAPPositiveDelimited", Err.Number, Err.Description, "Review: " & reviewNumber
End Sub


' ============================================================================
' HELPER FUNCTIONS
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: SortIndividuals
' ----------------------------------------------------------------------------
' PURPOSE:
'   Sorts the individual rows so that:
'   1. Head of household (relationship code "01") is first
'   2. Remaining members sorted by line number ascending
'
' WHY THIS MATTERS:
'   The schedule expects household members in a specific order. The head
'   of household should always be listed first, followed by other members
'   in line number order.
'
' PARAMETERS:
'   wsInd        - The BIS Individual worksheet
'   startRow     - First row of individuals for this case
'   stopRow      - Last row of individuals for this case
'   lastCol      - Last column of data
' ----------------------------------------------------------------------------
Private Sub SortIndividuals(ByRef wsInd As Worksheet, _
                            ByVal startRow As Long, _
                            ByVal stopRow As Long, _
                            ByVal lastCol As Long)
    On Error Resume Next
    
    ' First sort: By relationship code (column T) descending
    ' This puts "X" (head of household) at the top
    wsInd.Sort.SortFields.Clear
    wsInd.Sort.SortFields.Add Key:=wsInd.Range("T" & startRow & ":T" & stopRow), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With wsInd.Sort
        .SetRange wsInd.Range(wsInd.Cells(startRow, 1), wsInd.Cells(stopRow, lastCol))
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ' Second sort: Remaining members (rows 2+) by line number (column L) ascending
    If stopRow > startRow Then
        wsInd.Sort.SortFields.Clear
        wsInd.Sort.SortFields.Add Key:=wsInd.Range("L" & (startRow + 1) & ":L" & stopRow), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        With wsInd.Sort
            .SetRange wsInd.Range(wsInd.Cells(startRow + 1, 1), wsInd.Cells(stopRow, lastCol))
            .Header = xlNo
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End If
    
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' Sub: PopulateCaseLevelData
' ----------------------------------------------------------------------------
' PURPOSE:
'   Populates the FS Workbook with case-level data from BIS Case sheet.
'   This includes: phone number, dates, certification period, LIHEAP, rent, SUA.
'
' BIS CASE COLUMNS USED:
'   AJ = Telephone Number
'   M  = Application Date (YYYYMMDD)
'   X,Y,Z = Most Recent Action Date (YY, MM, DD)
'   AA,AB,AC = Recert Thru Date (YY, MM, DD)
'   AK = Coupon Amount (allotment)
'   AW = LIHEAP indicator (Y/N)
'   AU = Rent amount
'   AX = Standard Utility code (N/U/H/L/P)
'   74-134 = Absent Relatives data (30-column blocks)
' ----------------------------------------------------------------------------
Private Sub PopulateCaseLevelData(ByRef wsFSWB As Worksheet, _
                                   ByRef wsSchedule As Worksheet, _
                                   ByRef wsBISCase As Worksheet, _
                                   ByVal caseRowNum As Long, _
                                   ByRef incomeFreq As Variant)
    On Error Resume Next
    
    Dim temp As Variant
    Dim timeZero As Date
    Dim applDate As Date
    Dim mostRecentActionDate As Date
    Dim recertFromDate As Date
    Dim recertThruDate As Date
    Dim j As Long
    Dim writeRowAbs As Long
    
    ' Time zero used for date comparisons (date that represents "no date")
    timeZero = DateSerial(0, 0, 1)
    
    ' ----- TELEPHONE NUMBER -----
    wsFSWB.Range("D16").Value = wsBISCase.Range("AJ" & caseRowNum).Value
    
    ' ----- APPLICATION DATE -----
    temp = wsBISCase.Range("M" & caseRowNum).Value
    If temp = 0 Or temp = "" Then
        applDate = timeZero
    Else
        ' Convert YYYYMMDD to Date
        applDate = DateSerial(Val(Left(temp, 4)), Val(Mid(temp, 5, 2)), Val(Right(temp, 2)))
    End If
    
    If applDate > timeZero Then
        wsFSWB.Range("G20").Value = applDate
    End If
    
    ' ----- ACTION AND CERTIFICATION DATES -----
    ' Most recent action date from columns X, Y, Z (YY, MM, DD)
    mostRecentActionDate = DateSerial( _
        2000 + Val(wsBISCase.Range("X" & caseRowNum).Value), _
        Val(wsBISCase.Range("Y" & caseRowNum).Value), _
        Val(wsBISCase.Range("Z" & caseRowNum).Value))
    
    ' Recert from date (same X, Y but day = 1)
    recertFromDate = DateSerial( _
        2000 + Val(wsBISCase.Range("X" & caseRowNum).Value), _
        Val(wsBISCase.Range("Y" & caseRowNum).Value), _
        1)
    
    ' Recert thru date from columns AA, AB, AC
    recertThruDate = DateSerial( _
        2000 + Val(wsBISCase.Range("AA" & caseRowNum).Value), _
        Val(wsBISCase.Range("AB" & caseRowNum).Value), _
        Val(wsBISCase.Range("AC" & caseRowNum).Value))
    
    ' Most Recent Action - use later of action date or application date
    If mostRecentActionDate > applDate Then
        If mostRecentActionDate > timeZero Then
            wsFSWB.Range("G22").Value = mostRecentActionDate
        End If
    Else
        If applDate > timeZero Then
            wsFSWB.Range("G22").Value = applDate
        End If
    End If
    
    ' Certification period dates
    If recertFromDate > timeZero Then
        wsFSWB.Range("H24").Value = recertFromDate
    End If
    If recertThruDate > timeZero Then
        wsFSWB.Range("H25").Value = recertThruDate
    End If
    
    ' ----- COUPON/ALLOTMENT AMOUNT -----
    wsFSWB.Range("R42").Value = wsBISCase.Range("AK" & caseRowNum).Value
    
    ' ----- LIHEAP INDICATOR -----
    Select Case UCase(wsBISCase.Range("AW" & caseRowNum).Value)
        Case "N"
            wsFSWB.Range("L1499").Value = "No"
        Case "Y"
            wsFSWB.Range("L1499").Value = "Yes"
    End Select
    
    ' ----- ABSENT RELATIVES (Section C of Workbook) -----
    ' Data in columns 74-134 in 30-column blocks
    writeRowAbs = 25
    For j = 74 To 134 Step 30
        If Trim(wsBISCase.Cells(caseRowNum, j).Value) <> "" Then
            writeRowAbs = writeRowAbs + 2
            ' Name (First Last)
            wsFSWB.Range("K" & writeRowAbs).Value = _
                wsBISCase.Cells(caseRowNum, j + 1).Value & " " & _
                wsBISCase.Cells(caseRowNum, j).Value
            ' Relationship
            wsFSWB.Range("S" & writeRowAbs).Value = "LRR to " & wsBISCase.Cells(caseRowNum, j + 9).Value
            ' Other fields
            wsFSWB.Range("V" & writeRowAbs).Value = wsBISCase.Cells(caseRowNum, j + 2).Value
            wsFSWB.Range("Z" & writeRowAbs).Value = wsBISCase.Cells(caseRowNum, j + 3).Value
            wsFSWB.Range("Z" & writeRowAbs + 1).Value = _
                wsBISCase.Cells(caseRowNum, j + 5).Value & ", " & wsBISCase.Cells(caseRowNum, j + 6).Value
        Else
            Exit For
        End If
    Next j
    
    ' ----- RENT (Shelter Deduction) -----
    If wsBISCase.Range("AU" & caseRowNum).Value <> 0 Then
        wsFSWB.Range("N1462").Value = wsBISCase.Range("AU" & caseRowNum).Value
    End If
    
    ' ----- STANDARD UTILITY ALLOWANCE -----
    Select Case UCase(wsBISCase.Range("AX" & caseRowNum).Value)
        Case "N", "U"
            wsFSWB.Range("O1489").Value = "Non Heating"
        Case "H"
            wsFSWB.Range("O1489").Value = "Heating"
        Case "L"
            wsFSWB.Range("O1489").Value = "Limited"
        Case "P"
            wsFSWB.Range("O1489").Value = "Telephone"
    End Select
    
    ' ----- RESIDENCY CHECK (Element 140) -----
    If UCase(wsBISCase.Range("AH" & caseRowNum).Value) = "PA" Then
        wsFSWB.Shapes("CB 56").OLEFormat.Object.Value = 1
    End If
    
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' Sub: PopulateIndividualData
' ----------------------------------------------------------------------------
' PURPOSE:
'   Loops through all individuals for the case and populates the FS Workbook
'   with person-level data. Also builds lists of line numbers for various
'   eligibility categories (citizenship, E&T status, income types, etc.).
'
' This is the largest section - it processes each household member and
' populates multiple sections of the workbook based on their characteristics.
' ----------------------------------------------------------------------------
Private Sub PopulateIndividualData(ByRef wsFSWB As Worksheet, _
                                    ByRef wsSchedule As Worksheet, _
                                    ByRef wsBISInd As Worksheet, _
                                    ByVal startRow As Long, _
                                    ByVal stopRow As Long, _
                                    ByVal reviewNumber As String, _
                                    ByRef incomeFreq As Variant)
    On Error Resume Next
    
    Dim i As Long
    Dim writeRow As Long
    Dim temp As Variant
    Dim householdSize As Long
    
    ' Lists of line numbers for various categories
    Dim birthdate As Boolean
    Dim school As String, schoolEM As String, schoolNM As String
    Dim citizen As String, noncitizenQ As String, noncitizenNOT As String
    Dim household As String, householdSNAP As String
    Dim allcitizen As Long
    Dim mandatory As String, Age As String, health As String
    Dim student As String, employment As String, Other As String
    Dim EligibleABAWD As String, ExemptABAWD As String
    Dim RSDI_LNS As String, VA_LNS As String, SSI_LNS As String
    Dim SSP_LNS As String, UC_LNS As String, WC_LNS As String, DI_LNS As String
    
    ' Initialize variables
    birthdate = True
    allcitizen = 1
    householdSize = 0
    writeRow = 10
    
    ' Initialize checkboxes to default values (will be changed if conditions met)
    wsFSWB.Shapes("CB 104").OLEFormat.Object.Value = 1  ' Element 151
    wsFSWB.Shapes("CB 125").OLEFormat.Object.Value = 1  ' Element 160
    wsFSWB.Shapes("CB 146").OLEFormat.Object.Value = 1  ' Element 161
    wsFSWB.Shapes("CB 185").OLEFormat.Object.Value = 1  ' Element 163
    wsFSWB.Shapes("CB 186").OLEFormat.Object.Value = 1  ' Element 163
    wsFSWB.Shapes("CB 214").OLEFormat.Object.Value = 1  ' Element 165
    wsFSWB.Shapes("CB 217").OLEFormat.Object.Value = 1  ' Element 165
    wsFSWB.Shapes("CB 243").OLEFormat.Object.Value = 1  ' Element 166
    
    ' ========================================================================
    ' LOOP THROUGH ALL INDIVIDUALS
    ' ========================================================================
    For i = startRow To stopRow
        
        ' ----- SECTION B: HOUSEHOLD COMPOSITION -----
        writeRow = writeRow + 1
        If writeRow < 23 Then  ' Only room for 12 members on workbook
            ' Line Number
            wsFSWB.Range("K" & writeRow).Value = FormatLineNumber(wsBISInd.Range("L" & i).Value)
            
            ' Full Name (First Middle Last Suffix)
            wsFSWB.Range("M" & writeRow).Value = _
                Trim(wsBISInd.Range("N" & i).Value) & " " & _
                Trim(wsBISInd.Range("P" & i).Value) & " " & _
                Trim(wsBISInd.Range("O" & i).Value) & " " & _
                Trim(wsBISInd.Range("Q" & i).Value)
            
            ' Individual Category
            wsFSWB.Range("W" & writeRow).Value = wsBISInd.Range("J" & i).Value
            
            ' Date of Birth (convert YYYYMMDD)
            temp = wsBISInd.Range("R" & i).Value
            If Len(temp) >= 8 Then
                wsFSWB.Range("Y" & writeRow).Value = _
                    DateSerial(Val(Left(temp, 4)), Val(Mid(temp, 5, 2)), Val(Right(temp, 2)))
            End If
            
            ' Age
            wsFSWB.Range("AB" & writeRow).Value = wsBISInd.Range("S" & i).Value
            
            ' Relationship
            wsFSWB.Range("AD" & writeRow).Value = wsBISInd.Range("T" & i).Value
            
            ' SSN
            wsFSWB.Range("AF" & writeRow).Value = wsBISInd.Range("U" & i).Value
            
            ' SNAP Participation
            If wsBISInd.Range("Y" & i).Value = "NM" Then
                wsFSWB.Range("AJ" & writeRow).Value = "No"
            Else
                wsFSWB.Range("AJ" & writeRow).Value = "Yes"
                householdSize = householdSize + 1
            End If
        End If
        
        ' ----- ELEMENT 110: BIRTHDATE CHECK -----
        If wsBISInd.Range("R" & i).Value = "" Or Val(wsBISInd.Range("R" & i).Value) = 0 Then
            birthdate = False
        End If
        
        ' ----- ELEMENT 111: SCHOOL STATUS (AGE 18+) -----
        If Val(wsBISInd.Range("S" & i).Value) >= 18 And Val(wsBISInd.Range("W" & i).Value) <> 0 Then
            school = AppendLineNumber(school, wsBISInd.Range("L" & i).Value)
            
            ' Also track by eligibility status
            If wsBISInd.Range("Y" & i).Value = "EM" Then
                schoolEM = AppendLineNumber(schoolEM, wsBISInd.Range("L" & i).Value)
            ElseIf wsBISInd.Range("Y" & i).Value = "NM" Then
                schoolNM = AppendLineNumber(schoolNM, wsBISInd.Range("L" & i).Value)
            End If
        End If
        
        ' ----- ELEMENT 130: CITIZENSHIP STATUS -----
        If Val(wsBISInd.Range("AF" & i).Value) = 1 Then
            citizen = AppendLineNumber(citizen, wsBISInd.Range("L" & i).Value)
        ElseIf Val(wsBISInd.Range("AF" & i).Value) <> 1 Then
            allcitizen = 0
            If Left(wsBISInd.Range("Y" & i).Value, 1) = "E" Then
                noncitizenQ = AppendLineNumber(noncitizenQ, wsBISInd.Range("L" & i).Value)
            ElseIf Left(wsBISInd.Range("Y" & i).Value, 1) = "N" Then
                noncitizenNOT = AppendLineNumber(noncitizenNOT, wsBISInd.Range("L" & i).Value)
            End If
        End If
        
        ' ----- ELEMENT 150: HOUSEHOLD COMPOSITION -----
        Dim eligStatus As String
        eligStatus = UCase(wsBISInd.Range("Y" & i).Value)
        
        If eligStatus = "EM" Or eligStatus = "EB" Or eligStatus = "EW" Or eligStatus = "NM" Then
            household = AppendLineNumber(household, wsBISInd.Range("L" & i).Value)
        End If
        
        If eligStatus = "EM" Or eligStatus = "EB" Or eligStatus = "EW" Then
            householdSNAP = AppendLineNumber(householdSNAP, wsBISInd.Range("L" & i).Value)
        End If
        
        ' Check for elderly/disabled
        If Val(wsBISInd.Range("AE" & i).Value) = 2 Or Val(wsBISInd.Range("AE" & i).Value) = 3 Then
            wsFSWB.Shapes("CB 285").OLEFormat.Object.Value = 1
        End If
        
        ' ----- ELEMENT 151: DISQUALIFICATION CHECK -----
        If eligStatus = "DS" Or eligStatus = "DF" Then
            wsFSWB.Shapes("CB 104").OLEFormat.Object.Value = 0
        End If
        
        ' ----- ELEMENT 160: E&T EXEMPTION STATUS -----
        Dim etpCode As Long
        etpCode = Val(wsBISInd.Range("AE" & i).Value)
        
        If etpCode <> 30 Then
            wsFSWB.Shapes("CB 125").OLEFormat.Object.Value = 0
        End If
        
        ' Categorize by ETP code
        Select Case etpCode
            Case 30, 40
                mandatory = AppendLineNumber(mandatory, wsBISInd.Range("L" & i).Value)
            Case 1, 2
                Age = AppendLineNumber(Age, wsBISInd.Range("L" & i).Value)
            Case 3
                health = AppendLineNumber(health, wsBISInd.Range("L" & i).Value)
            Case 17
                employment = AppendLineNumber(employment, wsBISInd.Range("L" & i).Value)
            Case 20
                student = AppendLineNumber(student, wsBISInd.Range("L" & i).Value)
            Case 13
                ' Special case: 16-17 year old in school
                If (Val(wsBISInd.Range("AB" & i).Value) = 16 Or Val(wsBISInd.Range("AB" & i).Value) = 17) _
                   And Val(wsBISInd.Range("W" & i).Value) <> 0 Then
                    student = AppendLineNumber(student, wsBISInd.Range("L" & i).Value)
                End If
            Case Else
                Other = AppendLineNumber(Other, wsBISInd.Range("L" & i).Value)
        End Select
        
        ' ----- ELEMENT 161: ABAWD STATUS -----
        If eligStatus = "EB" Then
            wsFSWB.Shapes("CB 146").OLEFormat.Object.Value = 0
            wsFSWB.Shapes("CB 185").OLEFormat.Object.Value = 0
            wsFSWB.Shapes("CB 214").OLEFormat.Object.Value = 0
            wsFSWB.Shapes("CB 243").OLEFormat.Object.Value = 0
        Else
            wsFSWB.Shapes("CB 317").OLEFormat.Object.Value = 1
            Select Case etpCode
                Case 4, 6, 10, 14, 15, 16, 17, 18, 19, 20, 21
                    ExemptABAWD = AppendLineNumber(ExemptABAWD, wsBISInd.Range("L" & i).Value)
                Case Else
                    EligibleABAWD = AppendLineNumber(EligibleABAWD, wsBISInd.Range("L" & i).Value)
            End Select
        End If
        
        ' ----- UNEARNED INCOME LINE NUMBERS (Elements 331-343) -----
        If Val(wsBISInd.Range("BX" & i).Value) > 0 Then  ' RSDI
            RSDI_LNS = AppendLineNumber(RSDI_LNS, wsBISInd.Range("L" & i).Value)
        End If
        If Val(wsBISInd.Range("BZ" & i).Value) > 0 Then  ' VA
            VA_LNS = AppendLineNumber(VA_LNS, wsBISInd.Range("L" & i).Value)
        End If
        If Val(wsBISInd.Range("CB" & i).Value) > 0 Then  ' SSI
            SSI_LNS = AppendLineNumber(SSI_LNS, wsBISInd.Range("L" & i).Value)
        End If
        If Val(wsBISInd.Range("CI" & i).Value) > 0 Then  ' SSP
            SSP_LNS = AppendLineNumber(SSP_LNS, wsBISInd.Range("L" & i).Value)
        End If
        If Val(wsBISInd.Range("CD" & i).Value) > 0 Then  ' UC
            UC_LNS = AppendLineNumber(UC_LNS, wsBISInd.Range("L" & i).Value)
        End If
        If Val(wsBISInd.Range("CF" & i).Value) > 0 Then  ' WC
            WC_LNS = AppendLineNumber(WC_LNS, wsBISInd.Range("L" & i).Value)
        End If
        If Val(wsBISInd.Range("CH" & i).Value) > 0 Then  ' Deemed Income
            DI_LNS = AppendLineNumber(DI_LNS, wsBISInd.Range("L" & i).Value)
        End If
        
        ' ----- ADDITIONAL CHECKBOX UPDATES -----
        ' Element 163
        Dim csCode As Long
        csCode = Val(wsBISInd.Range("AC" & i).Value)
        If csCode = 21 Or csCode = 22 Or csCode = 23 Then
            wsFSWB.Shapes("CB 186").OLEFormat.Object.Value = 0
        End If
        
        ' Element 165
        If Val(wsBISInd.Range("AN" & i).Value) <> 0 Then
            wsFSWB.Shapes("CB 217").OLEFormat.Object.Value = 0
        End If
        If Val(wsBISInd.Range("AN" & i).Value) = 1 Or Val(wsBISInd.Range("AN" & i).Value) = 3 Then
            wsFSWB.Shapes("CB 223").OLEFormat.Object.Value = 1
        End If
        If Val(wsBISInd.Range("AN" & i).Value) = 2 Or Val(wsBISInd.Range("AN" & i).Value) = 4 Then
            wsFSWB.Shapes("CB 240").OLEFormat.Object.Value = 1
        End If
        
        ' Element 166
        Dim empCode As Long
        empCode = Val(wsBISInd.Range("AO" & i).Value)
        If empCode = 5 Or empCode = 6 Or empCode = 7 Or empCode = 10 Or empCode = 13 Then
            wsFSWB.Shapes("CB 225").OLEFormat.Object.Value = 1
        End If
        
        Dim ageCode As Long
        ageCode = Val(wsBISInd.Range("AB" & i).Value)
        If ageCode = 21 Or ageCode = 22 Or ageCode = 23 Then
            wsFSWB.Shapes("CB 246").OLEFormat.Object.Value = 1
        End If
        
        ' Element 170
        If Val(wsBISInd.Range("AC" & i).Value) = 24 Then
            wsFSWB.Shapes("CB 230").OLEFormat.Object.Value = 1
        End If
        
        ' Element 342 (Child Support)
        If Val(wsBISInd.Range("BI" & i).Value) > 0 Then
            wsFSWB.Shapes("CB 746").OLEFormat.Object.Value = 1
        End If
        
    Next i  ' End individual loop
    
    ' ========================================================================
    ' WRITE COLLECTED DATA TO WORKBOOK
    ' ========================================================================
    
    ' Element 110 - Birthdate check
    If birthdate = True Then
        wsFSWB.Shapes("CB 25").OLEFormat.Object.Value = 1
    End If
    
    ' Element 111 - School status
    wsFSWB.Range("I110").Value = school
    wsFSWB.Range("I112").Value = schoolEM
    wsFSWB.Range("I114").Value = schoolNM
    
    ' Element 130 - Citizenship
    Dim writeRowCitizen As Long
    writeRowCitizen = 121
    
    If allcitizen = 1 Then
        writeRowCitizen = writeRowCitizen + 1
        wsFSWB.Range("G" & writeRowCitizen).Value = "All"
        wsFSWB.Range("I" & writeRowCitizen).Value = "Citizen"
    Else
        If Len(citizen) > 0 Then
            writeRowCitizen = writeRowCitizen + 1
            wsFSWB.Range("G" & writeRowCitizen).Value = citizen
            wsFSWB.Range("I" & writeRowCitizen).Value = "Citizen"
        End If
        If Len(noncitizenQ) > 0 Then
            writeRowCitizen = writeRowCitizen + 1
            wsFSWB.Range("G" & writeRowCitizen).Value = noncitizenQ
            wsFSWB.Range("I" & writeRowCitizen).Value = "Non-Citizen"
        End If
        If Len(noncitizenNOT) > 0 Then
            writeRowCitizen = writeRowCitizen + 1
            wsFSWB.Range("G" & writeRowCitizen).Value = noncitizenNOT
            wsFSWB.Range("I" & writeRowCitizen).Value = "Non-Citizen"
        End If
    End If
    
    ' Element 150 - Household composition
    wsFSWB.Range("I213").Value = household
    wsFSWB.Range("I215").Value = householdSNAP
    
    ' Element 160 - E&T status
    Call WriteETPData(wsFSWB, mandatory, Age, health, student, employment, Other)
    
    ' Element 161 - ABAWD
    If wsFSWB.Shapes("CB 146").OLEFormat.Object.Value = 0 Then
        wsFSWB.Range("M344").Value = EligibleABAWD
        wsFSWB.Range("M348").Value = ExemptABAWD
        
        ' Element 163
        If Len(EligibleABAWD) > 0 And Len(ExemptABAWD) > 0 Then
            wsFSWB.Range("N370").Value = EligibleABAWD & ", " & ExemptABAWD
        ElseIf Len(EligibleABAWD) > 0 Then
            wsFSWB.Range("N370").Value = EligibleABAWD
        ElseIf Len(ExemptABAWD) > 0 Then
            wsFSWB.Range("N370").Value = ExemptABAWD
        End If
    End If
    
    ' Unearned income line numbers
    If Len(RSDI_LNS) > 0 Then wsFSWB.Range("K1087").Value = RSDI_LNS
    If Len(VA_LNS) > 0 Then wsFSWB.Range("N1096").Value = VA_LNS
    If Len(SSI_LNS) > 0 Then wsFSWB.Range("J1109").Value = SSI_LNS
    If Len(SSP_LNS) > 0 Then wsFSWB.Range("J1113").Value = SSP_LNS
    If Len(UC_LNS) > 0 Then wsFSWB.Range("N1124").Value = UC_LNS
    If Len(WC_LNS) > 0 Then wsFSWB.Range("N1198").Value = WC_LNS
    If Len(DI_LNS) > 0 Then wsFSWB.Range("N1235").Value = DI_LNS
    
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' Sub: WriteETPData
' ----------------------------------------------------------------------------
' PURPOSE:
'   Writes the E&T (Employment & Training) line number data to the workbook.
'   Element 160 requires listing household members by E&T category.
' ----------------------------------------------------------------------------
Private Sub WriteETPData(ByRef wsFSWB As Worksheet, _
                          ByVal mandatory As String, _
                          ByVal Age As String, _
                          ByVal health As String, _
                          ByVal student As String, _
                          ByVal employment As String, _
                          ByVal Other As String)
    
    Dim lnarray As Variant, exarray As Variant
    Dim writeRowETP As Long
    
    lnarray = Array(267, 269, 271, 273, 275, 276)
    exarray = Array(268, 270, 272, 274, 275, 276)
    
    writeRowETP = -1
    
    If Len(mandatory) > 0 Then
        writeRowETP = writeRowETP + 1
        wsFSWB.Range("G" & lnarray(writeRowETP)).Value = mandatory
        wsFSWB.Range("L" & exarray(writeRowETP)).Value = "mandatory"
    End If
    If Len(Age) > 0 Then
        writeRowETP = writeRowETP + 1
        wsFSWB.Range("G" & lnarray(writeRowETP)).Value = Age
        wsFSWB.Range("L" & exarray(writeRowETP)).Value = "age"
    End If
    If Len(health) > 0 Then
        writeRowETP = writeRowETP + 1
        wsFSWB.Range("G" & lnarray(writeRowETP)).Value = health
        wsFSWB.Range("L" & exarray(writeRowETP)).Value = "health"
    End If
    If Len(student) > 0 Then
        writeRowETP = writeRowETP + 1
        wsFSWB.Range("G" & lnarray(writeRowETP)).Value = student
        wsFSWB.Range("L" & exarray(writeRowETP)).Value = "student"
    End If
    If Len(employment) > 0 Then
        writeRowETP = writeRowETP + 1
        wsFSWB.Range("G" & lnarray(writeRowETP)).Value = employment
        wsFSWB.Range("L" & exarray(writeRowETP)).Value = "employment"
    End If
    If Len(Other) > 0 Then
        writeRowETP = writeRowETP + 1
        wsFSWB.Range("G" & lnarray(writeRowETP)).Value = Other
        wsFSWB.Range("L" & exarray(writeRowETP)).Value = "other"
    End If
End Sub

' ----------------------------------------------------------------------------
' Sub: PopulateScheduleSection4
' ----------------------------------------------------------------------------
' PURPOSE:
'   Populates Section 4 of the SNAP schedule (Household Characteristics).
'   This section has one row per household member with demographic data.
' ----------------------------------------------------------------------------
Private Sub PopulateScheduleSection4(ByRef wsSchedule As Worksheet, _
                                      ByRef wsBISInd As Worksheet, _
                                      ByVal startRow As Long, _
                                      ByVal stopRow As Long)
    On Error Resume Next
    
    Dim i As Long, j As Long
    Dim relCode As String
    Dim citizenCode As Long
    Dim raceCode As Long
    Dim eduLevel As Long
    
    i = startRow - 1
    
    ' Schedule Section 4 rows: 89, 92, 95... (every 3 rows) up to row 122
    For j = 89 To 122 Step 3
        i = i + 1
        If i > stopRow Then Exit For
        
        ' Line Number
        wsSchedule.Range("B" & j).Value = FormatLineNumber(wsBISInd.Range("L" & i).Value)
        
        ' SNAP Participation (01 = participating)
        If Left(wsBISInd.Range("Y" & i).Value, 1) = "E" Then
            wsSchedule.Range("E" & j).Value = "01"
        End If
        
        ' Relationship to Head of Household
        relCode = UCase(Trim(wsBISInd.Range("T" & i).Value))
        Select Case relCode
            Case "X", "X  "
                wsSchedule.Range("H" & j).Value = 1  ' HOH
            Case "W", "W  ", "H", "H  ", "CLH", "CLW"
                wsSchedule.Range("H" & j).Value = 2  ' Spouse
            Case "F", "F  ", "M", "M  ", "SF", "SF ", "SM", "SM "
                wsSchedule.Range("H" & j).Value = 3  ' Parent
            Case "D", "D  ", "SD", "SD ", "S", "S  ", "SS", "SS "
                wsSchedule.Range("H" & j).Value = 4  ' Child
            Case "NR", "NR "
                wsSchedule.Range("H" & j).Value = 6  ' Unrelated
            Case Else
                wsSchedule.Range("H" & j).Value = 5  ' Other related
        End Select
        
        ' Age
        wsSchedule.Range("J" & j).Value = wsBISInd.Range("S" & i).Value
        
        ' Gender
        Select Case UCase(wsBISInd.Range("CO" & i).Value)
            Case "F"
                wsSchedule.Range("M" & j).Value = "02"
            Case "M"
                wsSchedule.Range("M" & j).Value = "01"
        End Select
        
        ' Race
        raceCode = Val(wsBISInd.Range("CP" & i).Value)
        Select Case raceCode
            Case 1: wsSchedule.Range("P" & j).Value = "05"
            Case 3: wsSchedule.Range("P" & j).Value = "03"
            Case 4: wsSchedule.Range("P" & j).Value = "04"
            Case 5: wsSchedule.Range("P" & j).Value = "07"
            Case 6: wsSchedule.Range("P" & j).Value = "12"
            Case 7: wsSchedule.Range("P" & j).Value = "06"
            Case 8: wsSchedule.Range("P" & j).Value = "99"
        End Select
        
        ' Citizenship
        citizenCode = Val(wsBISInd.Range("AF" & i).Value)
        Select Case citizenCode
            Case 1
                wsSchedule.Range("S" & j).Value = "01"  ' US Born
            Case 4
                wsSchedule.Range("S" & j).Value = "05"  ' Refugee
        End Select
        
        ' Education Level
        eduLevel = Val(wsBISInd.Range("V" & i).Value)
        If eduLevel = 98 Then eduLevel = 0
        If eduLevel = 16 Then eduLevel = 14
        wsSchedule.Range("V" & j).Value = Format(eduLevel, "00")
        
    Next j
    
    On Error GoTo 0
End Sub


' ============================================================================
' UTILITY FUNCTIONS
' ============================================================================

' ----------------------------------------------------------------------------
' Function: FormatLineNumber
' ----------------------------------------------------------------------------
' PURPOSE:
'   Formats a line number as a 2-digit string with leading zero.
'   Example: 1 -> "01", 12 -> "12"
' ----------------------------------------------------------------------------
Private Function FormatLineNumber(ByVal lineNum As Variant) As String
    On Error Resume Next
    FormatLineNumber = Application.WorksheetFunction.Text(Val(lineNum), "00")
    On Error GoTo 0
End Function

' ----------------------------------------------------------------------------
' Function: AppendLineNumber
' ----------------------------------------------------------------------------
' PURPOSE:
'   Appends a line number to a comma-separated list.
'   Handles the first item (no comma) vs subsequent items (with comma).
' ----------------------------------------------------------------------------
Private Function AppendLineNumber(ByVal currentList As String, _
                                   ByVal lineNum As Variant) As String
    Dim formattedLN As String
    formattedLN = FormatLineNumber(lineNum)
    
    If Len(currentList) > 0 Then
        AppendLineNumber = currentList & ", " & formattedLN
    Else
        AppendLineNumber = formattedLN
    End If
End Function


