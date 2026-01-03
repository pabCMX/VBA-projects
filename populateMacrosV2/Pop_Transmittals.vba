Attribute VB_Name = "Pop_Transmittals"
' ============================================================================
' Pop_Transmittals - Transmittal Sheet Generation Module
' ============================================================================
' WHAT THIS MODULE DOES:
'   This module creates transmittal sheets (cover sheets) for batches of
'   QC reviews. Transmittals are sent to CAOs (County Assistance Offices)
'   along with case documentation requests.
'
' WHAT A TRANSMITTAL IS:
'   A transmittal is a cover sheet that accompanies a QC review request.
'   It includes:
'   - County and district information
'   - Client name and case/review numbers
'   - Clerk/supervisor contact information
'   - Space for tracking document receipt
'
' HOW IT WORKS:
'   1. User selects program and month from the main form
'   2. User selects the File of Records (list of cases)
'   3. System creates a new workbook for transmittals
'   4. For each case matching the program/month:
'      - Copies the transmittal template
'      - Fills in case-specific information
'      - Names the sheet with the review number
'
' FILE OF RECORDS STRUCTURE:
'   - Column A: Review Number
'   - Column B: Sample Month (YYYYMM format)
'   - Column D: County Number
'   - Column F: Case Number
'   - Column H: Last Name
'   - Column I: First Name
'
' PROGRAM REVIEW NUMBER RANGES:
'   - GA: 90-90
'   - MA Positive: 20-23
'   - FS Positive: 50-51
'   - FS Supplemental: 55-55
'   - FS Negative: 60-66
'   - TANF: 14-14
'   - MA Negative: 80-82
'
' DEPENDENCIES:
'   - Pop_Main: For program and month values from user form
'   - Populate sheet: For county lookup tables
'
' CHANGE LOG:
'   2026-01-03  Created from TransPopulate.vba
'               Added Option Explicit, V2 comments, error handling
' ============================================================================

Option Explicit

' ============================================================================
' MAIN TRANSMITTAL FUNCTION
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: Transmittal
' ----------------------------------------------------------------------------
' PURPOSE:
'   Main entry point for generating transmittal sheets.
'   Called from the main population form when "Transmittal" option is selected.
'
' STEPS:
'   1. Prompt user for File of Records
'   2. Create new workbook for transmittals
'   3. Copy template for each matching case
'   4. Fill in case-specific data
' ----------------------------------------------------------------------------
Public Sub Transmittal()
    On Error GoTo ErrorHandler
    
    ' ========================================================================
    ' VARIABLE DECLARATIONS
    ' ========================================================================
    
    Dim program As String
    Dim reviewMonth As Long
    Dim inputFile As Variant
    Dim fileName As String
    Dim monthStr As String
    
    Dim startRange As Long
    Dim endRange As Long
    Dim maxRow As Long
    Dim i As Long, j As Long
    Dim flag As Long
    
    Dim tempStr As String
    Dim tempStr2 As String
    Dim testArr As Variant
    Dim upper As Long
    Dim file_name As String
    
    ' ========================================================================
    ' 1. GET USER INPUT
    ' ========================================================================
    
    ' Prompt for File of Records
    inputFile = Application.GetOpenFilename( _
        "File of Records (*.xlsx), *.xlsx", , _
        "Select File of Records")
    
    If inputFile = False Then Exit Sub
    
    Application.ScreenUpdating = False
    
    ' Get program and month from the Populate sheet
    program = Worksheets("Populate").Cells(7, 23).Value
    reviewMonth = Worksheets("Populate").Cells(7, 26).Value
    
    ' ========================================================================
    ' 2. CREATE OUTPUT WORKBOOK
    ' ========================================================================
    
    monthStr = Application.WorksheetFunction.Text(reviewMonth, "MMMM YYYY")
    fileName = "Transmittals for " & program & " " & monthStr & ".xlsx"
    
    Application.DisplayAlerts = False
    Workbooks.Add
    ActiveWorkbook.SaveAs fileName:=fileName, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    
    ' ========================================================================
    ' 3. OPEN FILE OF RECORDS AND COPY DATA
    ' ========================================================================
    
    Workbooks.Open fileName:=inputFile, UpdateLinks:=False
    
    ' Copy the appropriate sheet based on program
    If program = "TANF" Or program = "GA" Or program = "FS Supplemental" Or _
       program = "FS Positive" Or program = "FS Negative" Then
        Sheets("FS Cash main file").Copy Before:=Workbooks(fileName).Sheets(2)
        Sheets("FS Cash main file").Name = "Temp"
    ElseIf program = "MA Negative" Or program = "MA Positive" Then
        Sheets("MA main file").Copy Before:=Workbooks(fileName).Sheets(2)
        Sheets("MA main file").Name = "Temp"
    Else
        Sheets("CAR main file").Copy Before:=Workbooks(fileName).Sheets(2)
        Sheets("CAR main file").Name = "Temp"
    End If
    
    ' Close File of Records
    testArr = Split(CStr(inputFile), "\")
    upper = UBound(testArr)
    file_name = testArr(upper)
    
    Application.DisplayAlerts = False
    Windows(file_name).Close
    Application.DisplayAlerts = True
    
    ' ========================================================================
    ' 4. FILTER BY PROGRAM AND MONTH
    ' ========================================================================
    
    Windows(fileName).Activate
    Sheets("Temp").Select
    
    ' Find max row
    maxRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, 1).End(xlUp).Row
    
    ' Check if month exists in file
    flag = 0
    For i = 2 To maxRow
        If Cells(i, 2).Value = reviewMonth Then
            flag = 1
            Exit For
        End If
    Next i
    
    If flag = 0 Then
        MsgBox reviewMonth & " Month not Found. Check File of Records.", vbExclamation
        Exit Sub
    End If
    
    ' Sort by review number
    Range("A2:II" & maxRow).Sort Key1:=Range("A2"), Order1:=xlAscending, _
        Header:=xlGuess, OrderCustom:=1, MatchCase:=False, _
        Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    ' Get review number range for this program
    Call GetProgramRange(program, startRange, endRange)
    
    ' Delete non-matching rows (work backwards)
    For i = maxRow To 1 Step -1
        Dim prefix As Long
        prefix = Val(Left(CStr(Cells(i, 1).Value), 2))
        
        If prefix < startRange Or prefix > endRange Or Cells(i, 2).Value <> reviewMonth Then
            Rows(i).Delete
        End If
    Next i
    
    ' Check if any cases remain
    If Range("A1").Value = "" Then
        MsgBox "No schedules found for selected program and month", vbExclamation
        Windows("populate.xlsm").Activate
        Sheets("Populate").Select
        Workbooks(fileName).Close False
        Exit Sub
    End If
    
    ' Recalculate max row
    maxRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, 1).End(xlUp).Row
    
    ' ========================================================================
    ' 5. CREATE TRANSMITTAL SHEETS
    ' ========================================================================
    
    ' Copy transmittal template
    Windows("populate.xlsm").Activate
    Sheets("Transmittal").Copy Before:=Workbooks(fileName).Sheets(1)
    
    ' Delete extra sheets
    Application.DisplayAlerts = False
    On Error Resume Next
    Sheets("Sheet1").Delete
    Sheets("Sheet2").Delete
    Sheets("Sheet3").Delete
    On Error GoTo ErrorHandler
    Application.DisplayAlerts = True
    
    ' Rename first sheet
    Sheets("Transmittal").Name = CStr(Sheets("Temp").Range("A1").Value)
    
    ' Create transmittal for each case
    For i = 1 To maxRow
        ' Save periodically to prevent memory issues
        If i Mod 50 = 0 Then
            Application.DisplayAlerts = False
            Workbooks(fileName).Save
            Application.DisplayAlerts = True
        End If
        
        If i > 1 Then
            ' Copy previous sheet as template
            tempStr = CStr(Sheets("Temp").Range("A" & i - 1).Value)
            Sheets(tempStr).Copy Before:=Sheets("Temp")
            tempStr = tempStr & " (2)"
            Sheets(tempStr).Name = CStr(Sheets("Temp").Range("A" & i).Value)
        End If
        
        ' Fill in county name
        Call FillCountyInfo(i, program)
        
    Next i
    
    ' Clean up
    Windows("populate.xlsm").Activate
    Sheets("Populate").Select
    
    Application.DisplayAlerts = False
    Workbooks(fileName).Activate
    Sheets("Temp").Delete
    Workbooks(fileName).Save
    Application.DisplayAlerts = True
    
    Application.ScreenUpdating = True
    
    MsgBox "Transmittals created successfully!", vbInformation
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Error in Transmittal: " & Err.Description, vbCritical
    LogError "Transmittal", Err.Number, Err.Description, "Program: " & program
End Sub


' ============================================================================
' HELPER FUNCTIONS
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: GetProgramRange
' ----------------------------------------------------------------------------
' PURPOSE:
'   Returns the valid review number prefix range for a program.
' ----------------------------------------------------------------------------
Private Sub GetProgramRange(ByVal program As String, _
                             ByRef startRange As Long, _
                             ByRef endRange As Long)
    Select Case program
        Case "GA"
            startRange = 90
            endRange = 90
        Case "MA Positive"
            startRange = 20
            endRange = 23
        Case "FS Positive"
            startRange = 50
            endRange = 51
        Case "FS Supplemental"
            startRange = 55
            endRange = 55
        Case "FS Negative"
            startRange = 60
            endRange = 66
        Case "TANF"
            startRange = 14
            endRange = 14
        Case "TANF CAR"
            startRange = 34
            endRange = 34
        Case "MA Negative"
            startRange = 80
            endRange = 82
        Case Else
            startRange = 0
            endRange = 0
    End Select
End Sub

' ----------------------------------------------------------------------------
' Sub: FillCountyInfo
' ----------------------------------------------------------------------------
' PURPOSE:
'   Fills in county and district information on the transmittal sheet.
'   Uses VLOOKUP to get county name from Populate sheet.
' ----------------------------------------------------------------------------
Private Sub FillCountyInfo(ByVal rowNum As Long, ByVal program As String)
    On Error Resume Next
    
    Dim tempStr As String
    Dim tempStr2 As String
    Dim countyNum As Long
    
    countyNum = Val(Sheets("Temp").Range("D" & rowNum).Value)
    
    ' County name lookup
    tempStr = "=VLOOKUP(Temp!R" & rowNum & "C4,[populate.xlsm]Populate!R2C30:R68C31,2,FALSE)"
    Range("C6").FormulaR1C1 = tempStr
    Range("C6").Copy
    Range("C6").PasteSpecial Paste:=xlPasteValues
    
    Range("C6").Value = Application.WorksheetFunction.Text(countyNum, "00") & " - " & _
                        Range("C6").Value & " CAO"
    
    ' District lookup (for counties with multiple districts)
    Range("C7").Value = ""
    Select Case countyNum
        Case 2
            tempStr2 = "=VLOOKUP(Temp!R" & rowNum & "C5,[populate.xlsm]Populate!R2C15:R9C18,4,FALSE)"
            Range("C7").FormulaR1C1 = tempStr2
        Case 9
            tempStr2 = "=VLOOKUP(Temp!R" & rowNum & "C5,[populate.xlsm]Populate!R10C15:R11C18,4,FALSE)"
            Range("C7").FormulaR1C1 = tempStr2
        Case 23
            tempStr2 = "=VLOOKUP(Temp!R" & rowNum & "C5,[populate.xlsm]Populate!R12C15:R13C18,4,FALSE)"
            Range("C7").FormulaR1C1 = tempStr2
        Case 40
            tempStr2 = "=VLOOKUP(Temp!R" & rowNum & "C5,[populate.xlsm]Populate!R14C15:R15C18,4,FALSE)"
            Range("C7").FormulaR1C1 = tempStr2
        Case 46
            tempStr2 = "=VLOOKUP(Temp!R" & rowNum & "C5,[populate.xlsm]Populate!R16C15:R17C18,4,FALSE)"
            Range("C7").FormulaR1C1 = tempStr2
        Case 51
            tempStr2 = "=VLOOKUP(Temp!R" & rowNum & "C5,[populate.xlsm]Populate!R18C15:R36C18,4,FALSE)"
            Range("C7").FormulaR1C1 = tempStr2
        Case 63
            tempStr2 = "=VLOOKUP(Temp!R" & rowNum & "C5,[populate.xlsm]Populate!R37C15:R38C18,4,FALSE)"
            Range("C7").FormulaR1C1 = tempStr2
        Case 65
            tempStr2 = "=VLOOKUP(Temp!R" & rowNum & "C5,[populate.xlsm]Populate!R39C15:R40C18,4,FALSE)"
            Range("C7").FormulaR1C1 = tempStr2
    End Select
    
    ' Append district if present
    If Range("C7").Value <> "" Then
        Range("C7").Copy
        Range("C7").PasteSpecial Paste:=xlPasteValues
        Range("C7").Value = Range("C7").Value & " District"
        Range("C6").Value = Range("C6").Value & " , " & Range("C7").Value
        Range("C7").Value = ""
    End If
    
    ' Client name (First Last)
    Range("B10").Value = Sheets("Temp").Range("I" & rowNum).Value & " " & _
                         Sheets("Temp").Range("H" & rowNum).Value
    
    ' Case / Review Number
    Range("G10").Value = Sheets("Temp").Range("F" & rowNum).Value & " / " & _
                         Sheets("Temp").Range("A" & rowNum).Value
    
    ' Clerk designation
    If program = "MA Positive" Or program = "MA Negative" Then
        Range("I17").Value = "Clerk"
    Else
        Range("I17").Value = "Clerk"
    End If
    
    Range("A2").Select
    
    On Error GoTo 0
End Sub


