Attribute VB_Name = "Review_GA_Elements"
' ============================================================================
' Review_GA_Elements - GA-Specific Utility Functions
' ============================================================================
' WHAT THIS MODULE DOES:
'   Contains utility functions and computation helpers specific to GA
'   (General Assistance) reviews. GA has unique proration calculations
'   and grant determination logic that differs from other programs.
'
' KEY FUNCTIONS:
'   - GAClear()          : Clear GA Computation sheet
'   - GAcompsheet()      : Calculate prorated amounts
'   - GAfinalresults()   : Transfer final results to schedule
'   - GAfinaldetermination() : Copy final determination values
'   - GAshow_form1/2()   : Display GA helper forms
'   - redisplayGAform1/2() : Redisplay forms
'
' GA PRORATION:
'   GA benefits are often prorated based on the number of days in the
'   benefit period. The GAcompsheet function implements a complex
'   proration formula based on lookup tables.
'
' CHANGE LOG:
'   2026-01-02  Refactored from GAGetElements.vba and Module1 -
'               consolidated GA-specific code, added Option Explicit
' ============================================================================

Option Explicit


' ============================================================================
' COMPUTATION SHEET FUNCTIONS
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: GAClear
' ----------------------------------------------------------------------------
' PURPOSE:
'   Clears the GA Computation sheet to prepare for new calculations.
'
' WHAT IT CLEARS:
'   - All numeric constants in columns A-L
'   - Name fields (B1, C2, E2)
'   - Resets line number labels
'   - Resets proration formulas in row 71
' ----------------------------------------------------------------------------
Public Sub GAClear()
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim i As Long
    Dim tempStr As String
    
    Set ws = ActiveWorkbook.Worksheets("GA Computation")
    
    ' Unprotect the sheet
    ws.Unprotect Password:=SHEET_PASSWORD
    
    ' Clear numeric constants
    ws.Range("A:L").SpecialCells(xlCellTypeConstants, 1).ClearContents
    
    ' Clear name fields
    ws.Range("B1").ClearContents
    ws.Range("C2").ClearContents
    ws.Range("E2").ClearContents
    
    ' Reset line number labels
    ws.Range("A7").Value = "1.  line              / "
    ws.Range("A8").Value = "2.  line              / "
    ws.Range("A9").Value = "3.  line              / "
    ws.Range("A15").Value = "7.  line              / "
    ws.Range("A16").Value = "8.  line              / "
    ws.Range("A17").Value = "9.  line              / "
    ws.Range("A18").Value = "10.  line              / "
    ws.Range("B76").Value = "Comments:"
    
    ' Reset proration formulas for columns 3-9
    For i = 3 To 9
        tempStr = "=IF(R52C" & i & ">R53C" & i & ",0," & _
                  "IF(AND(R69C" & i & "="""",R70C" & i & "=""""),""""," & _
                  "IF(R69C" & i & "="""",R70C" & i & ",R69C" & i & ")))"
        ws.Cells(71, i).FormulaR1C1 = tempStr
    Next i
    
    ' Reset proration formulas for columns 11-12
    For i = 11 To 12
        tempStr = "=IF(R52C" & i & ">R53C" & i & ",0," & _
                  "IF(AND(R69C" & i & "="""",R70C" & i & "=""""),""""," & _
                  "IF(R69C" & i & "="""",R70C" & i & ",R69C" & i & ")))"
        ws.Cells(71, i).FormulaR1C1 = tempStr
    Next i
    
    ' Re-protect the sheet
    ws.Protect Password:=SHEET_PASSWORD
    
    ws.Range("A1").Select
    
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' Sub: GAcompsheet
' ----------------------------------------------------------------------------
' PURPOSE:
'   Calculates prorated GA amounts based on the number of days in the
'   benefit period. Uses a complex formula that varies by day count.
'
' PRORATION LOGIC:
'   - Row 65 contains the number of days to prorate
'   - Row 63 contains the base amount
'   - Row 68 gets the calculated prorated amount
'   - Different formulas apply for 1-14 days
'
' NOTE:
'   Proration tables only go up to 14 days. For longer periods,
'   a supervisor should be consulted.
' ----------------------------------------------------------------------------
Public Sub GAcompsheet()
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim i As Long
    Dim hundInt As Long
    Dim baseAmount As Double
    Dim days As Long
    
    Set ws = ActiveSheet
    
    ' Process columns 3-14 (skip column 10)
    For i = 3 To 14
        If i <> 10 Then
            If ws.Cells(65, i).Value > 0 Then
                
                baseAmount = ws.Cells(63, i).Value
                days = ws.Cells(65, i).Value
                
                ' Calculate hundreds interval for formula
                If baseAmount < 101 Then
                    hundInt = Application.WorksheetFunction.RoundDown((baseAmount / 100) - 1, 0)
                Else
                    hundInt = Application.WorksheetFunction.RoundDown((baseAmount / 100), 0)
                End If
                
                Application.EnableEvents = False
                
                If days > 14 Then
                    MsgBox "Proration tables only go up to 14 days. " & _
                           "Please ask your Supervisor how to prorate beyond 14 days.", _
                           vbExclamation, "Proration Limit"
                Else
                    ' Apply proration formula based on number of days
                    ws.Cells(68, i).Value = CalculateProration(baseAmount, days, hundInt)
                End If
                
                Application.EnableEvents = True
            End If
        End If
    Next i
    
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' Function: CalculateProration
' ----------------------------------------------------------------------------
' PURPOSE:
'   Calculates the prorated amount based on the proration formula.
'   Each day count has a different multiplier and formula.
'
' PARAMETERS:
'   baseAmount - The base grant amount
'   days       - Number of days to prorate
'   hundInt    - The hundreds interval value
'
' RETURNS:
'   Double - The prorated amount
' ----------------------------------------------------------------------------
Private Function CalculateProration(ByVal baseAmount As Double, _
                                     ByVal days As Long, _
                                     ByVal hundInt As Long) As Double
    Dim remainder As Double
    remainder = baseAmount - 100 * hundInt
    
    Select Case days
        Case 1
            CalculateProration = 6.6 * hundInt + Round(0.0000472884 + 0.065710554 * remainder, 1)
        Case 2
            CalculateProration = 13.1 * hundInt + Round(-0.00017559 + 0.131430335 * remainder, 1)
        Case 3
            If baseAmount = 87.5 Then
                CalculateProration = 17.2
            Else
                CalculateProration = 19.7 * hundInt + Round(-0.00010194 + 0.197144297 * remainder, 1)
            End If
        Case 4
            If baseAmount = 5.9 Then
                CalculateProration = 1.5
            Else
                CalculateProration = 26.3 * hundInt + Round(-0.000130429 + 0.262858301 * remainder, 1)
            End If
        Case 5
            If baseAmount = 6.9 Then
                CalculateProration = 2.2
            Else
                CalculateProration = 32.9 * hundInt + Round(-0.001141619 + 0.328581779 * remainder, 1)
            End If
        Case 6
            If baseAmount = 100.3 Then
                CalculateProration = 39.6
            Else
                CalculateProration = 39.4 * hundInt + Round(-0.00018061 + 0.3942901 * remainder, 1)
            End If
        Case 7
            If baseAmount = 50.3 Then
                CalculateProration = 23.2
            Else
                CalculateProration = 46 * hundInt + Round(0.001345477 + 0.459995182 * remainder, 1)
            End If
        Case 8
            CalculateProration = 52.6 * hundInt + Round(-0.0000486654 + 0.5257144 * remainder, 1)
        Case 9
            CalculateProration = 59.1 * hundInt + Round(-0.000156025 + 0.591429951 * remainder, 1)
        Case 10
            CalculateProration = 65.7 * hundInt + Round(0.000167077 + 0.657139017 * remainder, 1)
        Case 11
            CalculateProration = 72.3 * hundInt + Round(-0.0000937807 + 0.722855619 * remainder, 1)
        Case 12
            If baseAmount = 46.6 Then
                CalculateProration = 36.8
            Else
                CalculateProration = 78.9 * hundInt + Round(0.0000808591 + 0.788573487 * remainder, 1)
            End If
        Case 13
            CalculateProration = 85.4 * hundInt + Round(0.0000748837 + 0.85428116 * remainder, 1)
        Case 14
            CalculateProration = 92 * hundInt + Round(0.000024456 + 0.91999952 * remainder, 1)
        Case Else
            CalculateProration = 0
    End Select
End Function


' ============================================================================
' RESULTS TRANSFER FUNCTIONS
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: GAfinalresults
' ----------------------------------------------------------------------------
' PURPOSE:
'   Transfers the final computed result to the appropriate cell in the
'   schedule. Reads the destination from cell AL77.
' ----------------------------------------------------------------------------
Public Sub GAfinalresults()
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim columnLetter As String
    Dim rowNumber As Long
    Dim destCell As String
    
    Set ws = ActiveSheet
    
    columnLetter = Right(ws.Range("AL77").Value, 1)
    rowNumber = 71
    
    destCell = columnLetter & rowNumber
    ws.Range(destCell).Value = ws.Range("K71").Value + ws.Range("L71").Value
    ws.Range(destCell).Select
    
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' Sub: GAfinaldetermination
' ----------------------------------------------------------------------------
' PURPOSE:
'   Copies the final determination values from the selected column to
'   the determination column (C). The source column is read from AL78.
' ----------------------------------------------------------------------------
Public Sub GAfinaldetermination()
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim srcCol As String
    Dim copyRanges() As Variant
    Dim i As Long
    
    Set ws = ActiveSheet
    
    srcCol = Right(ws.Range("AL78").Value, 1)
    
    ' Define the rows to copy
    copyRanges = Array("7:11", "15:18", "21", "25", "27", "29", "32", "35", _
                       "49", "54", "65", "67", "68", "73")
    
    For i = LBound(copyRanges) To UBound(copyRanges)
        ws.Range("C" & copyRanges(i)).Value = ws.Range(srcCol & copyRanges(i)).Value
    Next i
    
    On Error GoTo 0
End Sub


' ============================================================================
' FORM DISPLAY FUNCTIONS
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: GAshow_form1
' ----------------------------------------------------------------------------
' PURPOSE:
'   Displays the first GA helper UserForm.
' ----------------------------------------------------------------------------
Public Sub GAshow_form1()
    On Error Resume Next
    UF_GA_Helper1.Show
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' Sub: GAshow_form2
' ----------------------------------------------------------------------------
' PURPOSE:
'   Displays the second GA helper UserForm.
' ----------------------------------------------------------------------------
Public Sub GAshow_form2()
    On Error Resume Next
    UF_GA_Helper2.Show
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' Sub: redisplayGAform1
' ----------------------------------------------------------------------------
' PURPOSE:
'   Redisplays the first GA helper UserForm.
' ----------------------------------------------------------------------------
Public Sub redisplayGAform1()
    On Error Resume Next
    UF_GA_Helper1.Show
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' Sub: redisplayGAform2
' ----------------------------------------------------------------------------
' PURPOSE:
'   Redisplays the second GA helper UserForm.
' ----------------------------------------------------------------------------
Public Sub redisplayGAform2()
    On Error Resume Next
    UF_GA_Helper2.Show
    On Error GoTo 0
End Sub


