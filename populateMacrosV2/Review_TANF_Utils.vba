Attribute VB_Name = "Review_TANF_Utils"
' ============================================================================
' Review_TANF_Utils - TANF-Specific Utility Functions
' ============================================================================
' WHAT THIS MODULE DOES:
'   Contains utility functions and computation helpers specific to TANF
'   (Temporary Assistance for Needy Families) reviews. These functions
'   handle TANF-specific calculations, form display, and data manipulation.
'
' MIGRATED FROM V1:
'   TANFmod.vba - Subs: tanf(), TANFclear(), TANFCompute(), show_form1(),
'                 redisplayform1(), TANFfinalresults(), show_form2(),
'                 redisplayform2(), finaldetermination()
'
' V1 -> V2 FUNCTION MAPPING:
'   tanf()              -> CalculateTANFProration()
'   TANFclear()         -> ClearTANFComputation()
'   TANFCompute()       -> ComputeTANFBenefit()
'   show_form1()        -> ShowTANFForm1()
'   redisplayform1()    -> RedisplayTANFForm1()
'   TANFfinalresults()  -> TANFFinalResults()
'   show_form2()        -> ShowTANFForm2()
'   redisplayform2()    -> RedisplayTANFForm2()
'   finaldetermination() -> TANFFinalDetermination()
'
' WHY A SEPARATE MODULE:
'   TANF has unique requirements that don't apply to other programs:
'   - Grant calculation formulas based on family size and income
'   - Proration tables for partial month benefits (1-14 days)
'   - Specific form helpers for computation sheets
'   These are kept separate to avoid cluttering the common utilities.
'
' ALSO USED BY:
'   - GA (shares proration calculation approach)
'
' CHANGE LOG:
'   2026-01-03  Refactored from TANFmod.vba - renamed functions, added
'               Option Explicit, improved error handling, V2 commenting style
' ============================================================================

Option Explicit


' ============================================================================
' PRORATION CALCULATION FUNCTIONS
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: CalculateTANFProration
' ----------------------------------------------------------------------------
' PURPOSE:
'   Calculates prorated TANF amounts based on the number of days in the
'   benefit period. Uses complex formulas that vary by day count (1-14).
'
' V1 EQUIVALENT: tanf()
'
' PRORATION LOGIC:
'   - Row 66 contains the base amount for proration
'   - Row 68 contains the number of days to prorate
'   - Row 70 receives the calculated prorated amount
'   - Different formulas apply for 1-14 days
'
' NOTE:
'   Proration tables only go up to 14 days. For longer periods,
'   a supervisor should be consulted.
' ----------------------------------------------------------------------------
Public Sub CalculateTANFProration()
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim i As Long
    Dim hundInt As Long
    Dim baseAmount As Double
    Dim days As Long
    
    Set ws = ActiveSheet
    
    ' Process columns 3-14 for each family member
    For i = 3 To 14
        If ws.Cells(68, i).Value > 0 Then
            
            baseAmount = ws.Cells(66, i).Value
            days = ws.Cells(68, i).Value
            
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
                ws.Cells(70, i).Value = CalculateTANFProrationAmount(baseAmount, days, hundInt)
            End If
            
            Application.EnableEvents = True
        End If
    Next i
    
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' Function: CalculateTANFProrationAmount
' ----------------------------------------------------------------------------
' PURPOSE:
'   Calculates the prorated amount based on the proration formula.
'   Each day count has a different multiplier and formula.
'
' PARAMETERS:
'   baseAmount - The base grant amount
'   days       - Number of days to prorate (1-14)
'   hundInt    - The hundreds interval value
'
' RETURNS:
'   Double - The prorated amount
' ----------------------------------------------------------------------------
Private Function CalculateTANFProrationAmount(ByVal baseAmount As Double, _
                                               ByVal days As Long, _
                                               ByVal hundInt As Long) As Double
    Dim remainder As Double
    remainder = baseAmount - 100 * hundInt
    
    Select Case days
        Case 1
            CalculateTANFProrationAmount = 6.6 * hundInt + Round(0.0000472884 + 0.065710554 * remainder, 1)
        Case 2
            CalculateTANFProrationAmount = 13.1 * hundInt + Round(-0.00017559 + 0.131430335 * remainder, 1)
        Case 3
            If baseAmount = 87.5 Then
                CalculateTANFProrationAmount = 17.2
            Else
                CalculateTANFProrationAmount = 19.7 * hundInt + Round(-0.00010194 + 0.197144297 * remainder, 1)
            End If
        Case 4
            If baseAmount = 5.9 Then
                CalculateTANFProrationAmount = 1.5
            Else
                CalculateTANFProrationAmount = 26.3 * hundInt + Round(-0.000130429 + 0.262858301 * remainder, 1)
            End If
        Case 5
            If baseAmount = 6.9 Then
                CalculateTANFProrationAmount = 2.2
            Else
                CalculateTANFProrationAmount = 32.9 * hundInt + Round(-0.001141619 + 0.328581779 * remainder, 1)
            End If
        Case 6
            If baseAmount = 100.3 Then
                CalculateTANFProrationAmount = 39.6
            Else
                CalculateTANFProrationAmount = 39.4 * hundInt + Round(-0.00018061 + 0.3942901 * remainder, 1)
            End If
        Case 7
            If baseAmount = 50.3 Then
                CalculateTANFProrationAmount = 23.2
            Else
                CalculateTANFProrationAmount = 46 * hundInt + Round(0.001345477 + 0.459995182 * remainder, 1)
            End If
        Case 8
            CalculateTANFProrationAmount = 52.6 * hundInt + Round(-0.0000486654 + 0.5257144 * remainder, 1)
        Case 9
            CalculateTANFProrationAmount = 59.1 * hundInt + Round(-0.000156025 + 0.591429951 * remainder, 1)
        Case 10
            CalculateTANFProrationAmount = 65.7 * hundInt + Round(0.000167077 + 0.657139017 * remainder, 1)
        Case 11
            CalculateTANFProrationAmount = 72.3 * hundInt + Round(-0.0000937807 + 0.722855619 * remainder, 1)
        Case 12
            If baseAmount = 46.6 Then
                CalculateTANFProrationAmount = 36.8
            Else
                CalculateTANFProrationAmount = 78.9 * hundInt + Round(0.0000808591 + 0.788573487 * remainder, 1)
            End If
        Case 13
            CalculateTANFProrationAmount = 85.4 * hundInt + Round(0.0000748837 + 0.85428116 * remainder, 1)
        Case 14
            CalculateTANFProrationAmount = 92 * hundInt + Round(0.000024456 + 0.91999952 * remainder, 1)
        Case Else
            CalculateTANFProrationAmount = 0
    End Select
End Function


' ============================================================================
' COMPUTATION SHEET FUNCTIONS
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: ClearTANFComputation
' ----------------------------------------------------------------------------
' PURPOSE:
'   Clears the TANF Computation sheet to prepare for new calculations.
'   Prompts user for confirmation before clearing.
'
' V1 EQUIVALENT: TANFclear()
'
' WHAT IT CLEARS:
'   - All numeric constants in columns A-N
'   - Resets line number labels in rows 6-8, 15-18
'   - Resets proration formulas in row 71
'   - Clears comments area
' ----------------------------------------------------------------------------
Public Sub ClearTANFComputation()
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim i As Long
    Dim tempStr As String
    Dim ans As VbMsgBoxResult
    
    ans = MsgBox("Are you sure that you want to clear the entire Computation sheet?", vbYesNo)
    If ans = vbNo Then Exit Sub
    
    Set ws = ThisWorkbook.Worksheets("TANF Computation")
    
    ' Unprotect the sheet
    ws.Unprotect Password:=SHEET_PASSWORD
    
    ' Clear numeric constants in columns A-N
    ws.Range("A:N").SpecialCells(xlCellTypeConstants, 1).ClearContents
    
    ' Reset line number labels
    ws.Range("A6").Value = "1.  line              / "
    ws.Range("A7").Value = "2.  line              / "
    ws.Range("A8").Value = "3.  line              / "
    ws.Range("A15").Value = "7.  line              / "
    ws.Range("A16").Value = "8.  line              / "
    ws.Range("A17").Value = "9.  line              / "
    ws.Range("A18").Value = "10.  line              / "
    ws.Range("B78").Value = "Comments:"
    
    ' Reset proration formulas for columns 3-11
    For i = 3 To 11
        tempStr = "=IF(R49C" & i & ">R50C" & i & ",0," & _
                  "IF(AND(R72C" & i & "=1,R69C" & i & "=""""),R70C" & i & "*0.75," & _
                  "IF(AND(R72C" & i & "=1,R70C" & i & "=""""),R69C" & i & "*0.75," & _
                  "IF(AND(R69C" & i & "="""",R70C" & i & "=""""),""""," & _
                  "IF(R69C" & i & "="""",R70C" & i & ",R69C" & i & ")))))"
        ws.Cells(71, i).FormulaR1C1 = tempStr
    Next i
    
    ' Reset proration formulas for columns 13-14
    For i = 13 To 14
        tempStr = "=IF(R49C" & i & ">R50C" & i & ",0," & _
                  "IF(AND(R72C" & i & "=1,R69C" & i & "=""""),R70C" & i & "*0.75," & _
                  "IF(AND(R72C" & i & "=1,R70C" & i & "=""""),R69C" & i & "*0.75," & _
                  "IF(AND(R69C" & i & "="""",R70C" & i & "=""""),""""," & _
                  "IF(R69C" & i & "="""",R70C" & i & ",R69C" & i & ")))))"
        ws.Cells(71, i).FormulaR1C1 = tempStr
    Next i
    
    ' Re-protect the sheet
    ws.Protect Password:=SHEET_PASSWORD
    
    ws.Range("A1").Select
    
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' Sub: ComputeTANFBenefit
' ----------------------------------------------------------------------------
' PURPOSE:
'   Navigates to the TANF computation section of the worksheet.
'
' V1 EQUIVALENT: TANFCompute()
' ----------------------------------------------------------------------------
Public Sub ComputeTANFBenefit()
    On Error Resume Next
    Application.GoTo reference:=ActiveSheet.Range("I1"), Scroll:=True
    Range("M6").Select
    On Error GoTo 0
End Sub


' ============================================================================
' RESULTS TRANSFER FUNCTIONS
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: TANFFinalResults
' ----------------------------------------------------------------------------
' PURPOSE:
'   Transfers the final computed result from columns M and N to the
'   appropriate cell in the schedule. Reads destination from cell AL77.
'
' V1 EQUIVALENT: TANFfinalresults()
' ----------------------------------------------------------------------------
Public Sub TANFFinalResults()
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim columnLetter As String
    Dim rowNumber As Long
    Dim destCell As String
    
    Set ws = ActiveSheet
    
    ' Get destination column from AL77
    columnLetter = Right(ws.Range("AL77").Value, 1)
    rowNumber = 71
    
    ' Calculate and write combined result
    destCell = columnLetter & rowNumber
    ws.Range(destCell).Value = ws.Range("M71").Value + ws.Range("N71").Value
    ws.Range(destCell).Select
    
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' Sub: TANFFinalDetermination
' ----------------------------------------------------------------------------
' PURPOSE:
'   Copies the final determination values from the selected column to
'   the determination column (C). The source column is read from AL78.
'
' V1 EQUIVALENT: finaldetermination()
'
' LOGIC:
'   If the source column has a formula in row 71, copy specific cell ranges.
'   Otherwise, just copy the final value.
' ----------------------------------------------------------------------------
Public Sub TANFFinalDetermination()
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim srcCol As String
    
    Set ws = ActiveSheet
    
    ' Get source column from AL78
    srcCol = Right(ws.Range("AL78").Value, 1)
    
    ' Check if row 71 has a formula (indicates full computation)
    If ws.Range(srcCol & "71").HasFormula Then
        ' Copy all computation data to column C
        ws.Range(srcCol & "6:" & srcCol & "9").Copy
        ws.Range("C6").PasteSpecial
        ws.Range(srcCol & "11").Copy
        ws.Range("C11").PasteSpecial
        ws.Range(srcCol & "15:" & srcCol & "20").Copy
        ws.Range("C15").PasteSpecial
        ws.Range(srcCol & "22").Copy
        ws.Range("C22").PasteSpecial
        ws.Range(srcCol & "24:" & srcCol & "26").Copy
        ws.Range("C24").PasteSpecial
        ws.Range(srcCol & "30").Copy
        ws.Range("C30").PasteSpecial
        ws.Range(srcCol & "44").Copy
        ws.Range("C44").PasteSpecial
        ws.Range(srcCol & "46").Copy
        ws.Range("C46").PasteSpecial
        ws.Range(srcCol & "57").Copy
        ws.Range("C57").PasteSpecial
        ws.Range(srcCol & "60").Copy
        ws.Range("C60").PasteSpecial
        ws.Range(srcCol & "68:" & srcCol & "69").Copy
        ws.Range("C68").PasteSpecial
        ws.Range(srcCol & "74").Copy
        ws.Range("C74").PasteSpecial
        ws.Range(srcCol & "76").Copy
        ws.Range("C76").PasteSpecial
        ws.Range(srcCol & "78").Copy
        ws.Range("C78").PasteSpecial
    Else
        ' Just copy the final value
        ws.Range("C71").Value = ws.Range(srcCol & "71").Value
    End If
    
    ws.Range("C76").Select
    
    On Error GoTo 0
End Sub


' ============================================================================
' FORM DISPLAY FUNCTIONS
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: ShowTANFForm1
' ----------------------------------------------------------------------------
' PURPOSE:
'   Displays the first TANF helper UserForm (UF_TANF_Helper).
'
' V1 EQUIVALENT: show_form1()
' ----------------------------------------------------------------------------
Public Sub ShowTANFForm1()
    On Error Resume Next
    UF_TANF_Helper.Show
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' Sub: RedisplayTANFForm1
' ----------------------------------------------------------------------------
' PURPOSE:
'   Redisplays the first TANF helper UserForm.
'
' V1 EQUIVALENT: redisplayform1()
' ----------------------------------------------------------------------------
Public Sub RedisplayTANFForm1()
    On Error Resume Next
    UF_TANF_Helper.Show
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' Sub: ShowTANFForm2
' ----------------------------------------------------------------------------
' PURPOSE:
'   Displays the second TANF helper UserForm (UF_TANF_Results).
'
' V1 EQUIVALENT: show_form2()
' ----------------------------------------------------------------------------
Public Sub ShowTANFForm2()
    On Error Resume Next
    UF_TANF_Results.Show
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' Sub: RedisplayTANFForm2
' ----------------------------------------------------------------------------
' PURPOSE:
'   Redisplays the second TANF helper UserForm.
'
' V1 EQUIVALENT: redisplayform2()
' ----------------------------------------------------------------------------
Public Sub RedisplayTANFForm2()
    On Error Resume Next
    UF_TANF_Results.Show
    On Error GoTo 0
End Sub


' ============================================================================
' RESULTS COLUMN FORM FUNCTIONS
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: ShowTANFResultsColumn
' ----------------------------------------------------------------------------
' PURPOSE:
'   Displays the TANF Results Column selection form.
' ----------------------------------------------------------------------------
Public Sub ShowTANFResultsColumn()
    On Error Resume Next
    UF_TANF_ResultsColumn.Show
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' Sub: ShowTANFFinalDetermination
' ----------------------------------------------------------------------------
' PURPOSE:
'   Displays the TANF Final Determination form.
' ----------------------------------------------------------------------------
Public Sub ShowTANFFinalDetermination()
    On Error Resume Next
    UF_TANF_FinalDetermination.Show
    On Error GoTo 0
End Sub


