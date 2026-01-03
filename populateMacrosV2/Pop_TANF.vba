Attribute VB_Name = "Pop_TANF"
' ============================================================================
' Pop_TANF - TANF Schedule Population Module
' ============================================================================
' WHAT THIS MODULE DOES:
'   This module reads data from BIS delimited files and populates TANF
'   (Temporary Assistance for Needy Families) review schedule templates.
'   TANF is the cash assistance program for families with children.
'
' HOW TANF DIFFERS FROM SNAP:
'   - Uses "TANF Workbook" instead of "FS Workbook"
'   - Different eligibility status codes (ES/EC instead of EM/EB)
'   - Different column mappings in BIS file (Age in T vs S, etc.)
'   - Grant calculation instead of allotment calculation
'   - Different element numbers and checkbox mappings
'
' BIS FILE STRUCTURE:
'   Same as SNAP Positive:
'   - "Case" sheet: One row per case with case-level data
'   - "Individual" sheet: One row per household member
'
' KEY CONCEPTS:
'   - AU (Assistance Unit): The group receiving TANF benefits
'   - Grant: Monthly cash payment amount
'   - Time Limits: Federal 60-month lifetime limit on TANF
'   - Work Requirements: Adults must participate in work activities
'
' DEPENDENCIES:
'   - Common_Utils: For helper functions
'   - Config_Settings: For income frequency multipliers
'   - Pop_Main: Sets up wb_bis and wb workbook references
'
' CHANGE LOG:
'   2026-01-03  Created from Populate_TANF_delimited_mod.vba
'               Added Option Explicit, V2 comments, error handling
' ============================================================================

Option Explicit

' ============================================================================
' MAIN POPULATION FUNCTION
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: PopulateTANFDelimited
' ----------------------------------------------------------------------------
' PURPOSE:
'   Main entry point for populating a TANF schedule from BIS data.
'   Very similar structure to SNAP Positive but with TANF-specific mappings.
'
' PRE-CONDITIONS:
'   - wb_bis contains the opened BIS delimited file
'   - wb contains the schedule workbook to populate
'   - program is set to "TANF"
' ----------------------------------------------------------------------------
Public Sub PopulateTANFDelimited()
    On Error GoTo ErrorHandler
    
    ' ========================================================================
    ' VARIABLE DECLARATIONS
    ' ========================================================================
    
    Dim wbSch As Workbook
    Dim wsBISCase As Worksheet
    Dim wsBISInd As Worksheet
    Dim wsTANFWorkbook As Worksheet
    Dim wsSchedule As Worksheet
    
    Dim reviewNumber As String
    Dim caseRow As Range
    Dim maxrowBISCase As Long
    Dim maxrowBISInd As Long
    Dim lastColBISInd As Long
    Dim bisIndStartRow As Long
    Dim bisIndStopRow As Long
    
    Dim incomeFreq As Variant
    incomeFreq = GetIncomeMultiplierArray()
    
    Dim ws As Worksheet
    Dim i As Long
    Dim temp As Variant
    Dim writeRow As Long
    Dim householdSize As Long
    
    ' ========================================================================
    ' 1. SETUP - GET WORKBOOK REFERENCES
    ' ========================================================================
    
    Set wbSch = wb
    
    ' Find review number from sheet name
    For Each ws In wbSch.Worksheets
        If Val(ws.Name) > 1000 Then
            reviewNumber = ws.Name
            Exit For
        End If
    Next ws
    
    If reviewNumber = "" Then
        MsgBox "Could not find schedule sheet", vbCritical
        Exit Sub
    End If
    
    Set wsBISCase = wb_bis.Worksheets("Case")
    Set wsBISInd = wb_bis.Worksheets("Individual")
    Set wsTANFWorkbook = wbSch.Worksheets("TANF Workbook")
    Set wsSchedule = wbSch.Worksheets(reviewNumber)
    
    ' ========================================================================
    ' 2. FIND CASE IN BIS FILE
    ' ========================================================================
    
    maxrowBISCase = wsBISCase.Cells(wsBISCase.Rows.Count, 1).End(xlUp).Row
    
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
    
    maxrowBISInd = wsBISInd.Cells(wsBISInd.Rows.Count, 1).End(xlUp).Row
    lastColBISInd = wsBISInd.Cells(1, wsBISInd.Columns.Count).End(xlToLeft).Column
    
    bisIndStartRow = 0
    bisIndStopRow = 0
    
    For i = 2 To maxrowBISInd
        If Val(wsBISInd.Range("C" & i).Value) = Val(reviewNumber) Then
            If bisIndStartRow = 0 Then bisIndStartRow = i
            bisIndStopRow = i
        ElseIf bisIndStartRow > 0 Then
            Exit For
        End If
    Next i
    
    ' ========================================================================
    ' 4. SORT INDIVIDUALS (HOH FIRST)
    ' ========================================================================
    
    If bisIndStopRow > bisIndStartRow Then
        Call SortTANFIndividuals(wsBISInd, bisIndStartRow, bisIndStopRow, lastColBISInd)
    End If
    
    ' ========================================================================
    ' 5. POPULATE CASE-LEVEL DATA
    ' ========================================================================
    
    ' Telephone Number (Column AB for TANF)
    wsTANFWorkbook.Range("D20").Value = wsBISCase.Range("AB" & caseRow.Row).Value
    
    ' ========================================================================
    ' 6. POPULATE INDIVIDUAL-LEVEL DATA
    ' ========================================================================
    
    If bisIndStartRow > 0 Then
        writeRow = 10
        householdSize = 0
        
        For i = bisIndStartRow To bisIndStopRow
            writeRow = writeRow + 1
            
            If writeRow < 23 Then  ' Room for 12 members
                ' Line Number
                wsTANFWorkbook.Range("J" & writeRow).Value = FormatLineNumber(wsBISInd.Range("L" & i).Value)
                
                ' Full Name
                wsTANFWorkbook.Range("L" & writeRow).Value = _
                    Trim(wsBISInd.Range("N" & i).Value) & " " & _
                    Trim(wsBISInd.Range("P" & i).Value) & " " & _
                    Trim(wsBISInd.Range("O" & i).Value) & " " & _
                    Trim(wsBISInd.Range("Q" & i).Value)
                
                ' Individual Category
                wsTANFWorkbook.Range("AC" & writeRow).Value = wsBISInd.Range("J" & i).Value
                
                ' Date of Birth
                temp = wsBISInd.Range("R" & i).Value
                If Len(temp) >= 8 Then
                    wsTANFWorkbook.Range("V" & writeRow).Value = _
                        DateSerial(Val(Left(temp, 4)), Val(Mid(temp, 5, 2)), Val(Right(temp, 2)))
                End If
                
                ' Age (Column T for TANF, was S for SNAP)
                wsTANFWorkbook.Range("Y" & writeRow).Value = wsBISInd.Range("T" & i).Value
                
                ' Relationship (Column X for TANF)
                wsTANFWorkbook.Range("AA" & writeRow).Value = wsBISInd.Range("X" & i).Value
                
                ' SSN (Column Z for TANF)
                wsTANFWorkbook.Range("AE" & writeRow).Value = wsBISInd.Range("Z" & i).Value
                
                ' TANF Participation (ES/EC = receiving)
                Dim eligStatus As String
                eligStatus = UCase(wsBISInd.Range("AD" & i).Value)
                
                If eligStatus = "ES" Or eligStatus = "EC" Then
                    wsTANFWorkbook.Range("AI" & writeRow).Value = "Yes"
                    wsTANFWorkbook.Range("AJ" & writeRow).Value = "Yes"
                Else
                    wsTANFWorkbook.Range("AI" & writeRow).Value = "No"
                    wsTANFWorkbook.Range("AJ" & writeRow).Value = "No"
                    householdSize = householdSize + 1
                End If
            End If
        Next i
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in PopulateTANFDelimited: " & Err.Description, vbCritical
    LogError "PopulateTANFDelimited", Err.Number, Err.Description, "Review: " & reviewNumber
End Sub


' ============================================================================
' HELPER FUNCTIONS
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: SortTANFIndividuals
' ----------------------------------------------------------------------------
' PURPOSE:
'   Sorts individuals for TANF. Uses Column X for relationship (HOH indicator)
'   and Column L for line number - slightly different than SNAP.
' ----------------------------------------------------------------------------
Private Sub SortTANFIndividuals(ByRef wsInd As Worksheet, _
                                 ByVal startRow As Long, _
                                 ByVal stopRow As Long, _
                                 ByVal lastCol As Long)
    On Error Resume Next
    
    ' Sort by relationship (Column X) descending to get HOH first
    wsInd.Sort.SortFields.Clear
    wsInd.Sort.SortFields.Add Key:=wsInd.Range("X" & startRow & ":X" & stopRow), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With wsInd.Sort
        .SetRange wsInd.Range(wsInd.Cells(startRow, 1), wsInd.Cells(stopRow, lastCol))
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ' Sort remaining by line number ascending
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
' Function: FormatLineNumber
' ----------------------------------------------------------------------------
Private Function FormatLineNumber(ByVal lineNum As Variant) As String
    On Error Resume Next
    FormatLineNumber = Application.WorksheetFunction.Text(Val(lineNum), "00")
    On Error GoTo 0
End Function


