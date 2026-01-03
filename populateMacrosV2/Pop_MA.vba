Attribute VB_Name = "Pop_MA"
' ============================================================================
' Pop_MA - Medical Assistance Schedule Population Module
' ============================================================================
' WHAT THIS MODULE DOES:
'   This module reads data from BIS delimited files and populates MA
'   (Medical Assistance / Medicaid) review schedule templates.
'
' HOW MA DIFFERS FROM SNAP/TANF:
'   - Uses "MA Workbook" worksheet
'   - Different eligibility categories (MAGI, ABD, etc.)
'   - Income calculations use MAGI methodology
'   - Different column mappings in BIS file
'   - Coverage groups and aid categories matter
'
' MA REVIEW TYPES:
'   - MA Positive: Approved MA cases being reviewed
'   - MA Negative: Denied/terminated MA cases
'   - MA PE: Presumptive Eligibility cases
'
' KEY CONCEPTS:
'   - MAGI: Modified Adjusted Gross Income (ACA methodology)
'   - FPL: Federal Poverty Level (eligibility threshold)
'   - Coverage Group: Type of Medicaid coverage
'   - Aid Category: Specific eligibility category
'
' BIS COLUMN DIFFERENCES:
'   - Phone: Column AB
'   - Dates: Columns AC-AH
'   - Individual data: Different column layout than SNAP
'
' DEPENDENCIES:
'   - Common_Utils: For helper functions
'   - Config_Settings: For constants
'   - Pop_Main: Sets up wb_bis and wb workbook references
'
' CHANGE LOG:
'   2026-01-03  Created from Populate_MA_delimited_mod.vba
'               Added Option Explicit, V2 comments, error handling
' ============================================================================

Option Explicit

' ============================================================================
' MAIN POPULATION FUNCTION
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: PopulateMADelimited
' ----------------------------------------------------------------------------
' PURPOSE:
'   Main entry point for populating an MA schedule from BIS data.
'   Handles both MA Positive and other MA review types.
' ----------------------------------------------------------------------------
Public Sub PopulateMADelimited()
    On Error GoTo ErrorHandler
    
    ' ========================================================================
    ' VARIABLE DECLARATIONS
    ' ========================================================================
    
    Dim wbSch As Workbook
    Dim wsBISCase As Worksheet
    Dim wsBISInd As Worksheet
    Dim wsMAWorkbook As Worksheet
    Dim wsSchedule As Worksheet
    
    Dim reviewNumber As String
    Dim caseRow As Range
    Dim maxrowBISCase As Long
    Dim maxrowBISInd As Long
    Dim lastColBISInd As Long
    Dim bisIndStartRow As Long
    Dim bisIndStopRow As Long
    
    Dim ws As Worksheet
    Dim i As Long
    Dim temp As Variant
    Dim timeZero As Date
    Dim openDate As Date
    Dim actionDate As Date
    
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
    Set wsMAWorkbook = wbSch.Worksheets("MA Workbook")
    Set wsSchedule = wbSch.Worksheets(reviewNumber)
    
    ' ========================================================================
    ' 2. FIND CASE IN BIS FILE
    ' ========================================================================
    
    maxrowBISCase = wsBISCase.Cells(wsBISCase.Rows.Count, 1).End(xlUp).Row
    lastColBISInd = wsBISCase.Cells(1, wsBISCase.Columns.Count).End(xlToLeft).Column
    
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
    ' 4. POPULATE CASE-LEVEL DATA (MA POSITIVE)
    ' ========================================================================
    
    If program = "MA Positive" Then
        timeZero = DateSerial(0, 0, 1)
        
        ' Telephone Number (Column AB)
        wsMAWorkbook.Range("D20").Value = wsBISCase.Range("AB" & caseRow.Row).Value
        
        ' Open Date (Columns AC, AD, AE = Year, Month, Day)
        openDate = DateSerial( _
            Val(wsBISCase.Range("AC" & caseRow.Row).Value), _
            Val(wsBISCase.Range("AD" & caseRow.Row).Value), _
            Val(wsBISCase.Range("AE" & caseRow.Row).Value))
        
        ' Action Date (Columns AF, AG, AH)
        actionDate = DateSerial( _
            Val(wsBISCase.Range("AF" & caseRow.Row).Value), _
            Val(wsBISCase.Range("AG" & caseRow.Row).Value), _
            Val(wsBISCase.Range("AH" & caseRow.Row).Value))
        
        ' Populate dates if valid
        ' (Note: Original code had date logic commented out - keeping structure)
    End If
    
    ' ========================================================================
    ' 5. POPULATE INDIVIDUAL-LEVEL DATA
    ' ========================================================================
    
    If bisIndStartRow > 0 Then
        Dim writeRow As Long
        writeRow = 10
        
        For i = bisIndStartRow To bisIndStopRow
            writeRow = writeRow + 1
            
            If writeRow < 23 Then  ' Room for 12 members
                ' Line Number
                wsMAWorkbook.Range("J" & writeRow).Value = FormatLineNumber(wsBISInd.Range("L" & i).Value)
                
                ' Full Name
                wsMAWorkbook.Range("L" & writeRow).Value = _
                    Trim(wsBISInd.Range("N" & i).Value) & " " & _
                    Trim(wsBISInd.Range("P" & i).Value) & " " & _
                    Trim(wsBISInd.Range("O" & i).Value) & " " & _
                    Trim(wsBISInd.Range("Q" & i).Value)
                
                ' Date of Birth
                temp = wsBISInd.Range("R" & i).Value
                If Len(temp) >= 8 Then
                    wsMAWorkbook.Range("V" & writeRow).Value = _
                        DateSerial(Val(Left(temp, 4)), Val(Mid(temp, 5, 2)), Val(Right(temp, 2)))
                End If
                
                ' Age
                wsMAWorkbook.Range("Y" & writeRow).Value = wsBISInd.Range("S" & i).Value
                
                ' Relationship
                wsMAWorkbook.Range("AA" & writeRow).Value = wsBISInd.Range("T" & i).Value
                
                ' SSN
                wsMAWorkbook.Range("AE" & writeRow).Value = wsBISInd.Range("U" & i).Value
            End If
        Next i
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in PopulateMADelimited: " & Err.Description, vbCritical
    LogError "PopulateMADelimited", Err.Number, Err.Description, "Review: " & reviewNumber
End Sub


' ============================================================================
' HELPER FUNCTIONS
' ============================================================================

' ----------------------------------------------------------------------------
' Function: FormatLineNumber
' ----------------------------------------------------------------------------
Private Function FormatLineNumber(ByVal lineNum As Variant) As String
    On Error Resume Next
    FormatLineNumber = Application.WorksheetFunction.Text(Val(lineNum), "00")
    On Error GoTo 0
End Function


