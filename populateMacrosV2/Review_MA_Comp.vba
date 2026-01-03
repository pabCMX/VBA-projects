Attribute VB_Name = "Review_MA_Comp"
' ============================================================================
' Review_MA_Comp - MA Computation Helpers
' ============================================================================
' WHAT THIS MODULE DOES:
'   Contains computation helper functions specific to MA (Medical Assistance)
'   Positive reviews. MA has unique eligibility calculations based on
'   Modified Adjusted Gross Income (MAGI) and other factors.
'
' MIGRATED FROM V1:
'   MA_Comp_mod.vba - Subs: MA_Comp_finalresults(), show_formMAC2(),
'                     redisplayformMAC2(), MA_Comp_finalresults3(), show_formMAC3(),
'                     redisplayformMAC3(), MAClearButtons1to15(), etc.,
'                     resources_ex(), ScheduleClean2(), TPL_Fill(), TPL_Clear(),
'                     MA_Comp_Transfer_SectionA(), MA_Comp_Clear_Wages(),
'                     MA_Comp_Clear_Income(), MA_Comp_Transfer_Workbook()
'
' V1 -> V2 FUNCTION MAPPING:
'   MA_Comp_finalresults()    -> MACompFinalResults()
'   MA_Comp_finalresults3()   -> MACompFinalResults3()
'   show_formMAC2()           -> ShowMACompForm2()
'   redisplayformMAC2()       -> RedisplayMACompForm2()
'   show_formMAC3()           -> ShowMACompForm3()
'   redisplayformMAC3()       -> RedisplayMACompForm3()
'   MA_Comp_Transfer_SectionA() -> TransferMASectionA()
'   MA_Comp_Clear_Wages()     -> ClearMAWages()
'   MA_Comp_Clear_Income()    -> ClearMAIncome()
'   MA_Comp_Transfer_Workbook() -> TransferMAToWorkbook()
'   MAClearButtons*()         -> ClearMAButtons()
'   ScheduleClean2()          -> PopulateMASchedule()
'   TPL_Fill()                -> FillTPLSection()
'   TPL_Clear()               -> ClearTPLSection()
'
' WHY A SEPARATE MODULE:
'   MA Positive reviews have complex calculation requirements:
'   - MAGI-based income calculations
'   - Household composition rules differ from other programs
'   - Multiple coverage groups with different rules
'
' CHANGE LOG:
'   2026-01-03  Refactored from MA_Comp_mod.vba - added Option Explicit,
'               improved organization, V2 commenting style
' ============================================================================

Option Explicit

' Module-level variable for column selection (used by UserForms)
Public MACompColumnLetter As String


' ============================================================================
' RESULTS TRANSFER FUNCTIONS
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: MACompFinalResults
' ----------------------------------------------------------------------------
' PURPOSE:
'   Transfers final computation results from the selected column to column C.
'   Used to copy QC-determined values to the final determination column.
'
' V1 EQUIVALENT: MA_Comp_finalresults()
' ----------------------------------------------------------------------------
Public Sub MACompFinalResults()
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim i As Long
    
    Set ws = ActiveSheet
    
    ' Copy main data
    ws.Range("C8").Value = ws.Range(MACompColumnLetter & "8").Value
    ws.Range("C144").Value = ws.Range(MACompColumnLetter & "144").Value
    
    ' Copy category and status code based on column
    Select Case MACompColumnLetter
        Case "D", "E", "F"
            ws.Range("C6").Value = ws.Range("E6").Value
            ws.Range("C7").Value = ws.Range("E7").Value
        Case "H", "I", "J"
            ws.Range("C6").Value = ws.Range("I6").Value
            ws.Range("C7").Value = ws.Range("I7").Value
        Case "K", "L", "M"
            ws.Range("C6").Value = ws.Range("L6").Value
            ws.Range("C7").Value = ws.Range("L7").Value
    End Select
    
    ' Copy data rows - grouped by section
    For i = 10 To 15: ws.Range("C" & i).Value = ws.Range(MACompColumnLetter & i).Value: Next i
    For i = 17 To 21: ws.Range("C" & i).Value = ws.Range(MACompColumnLetter & i).Value: Next i
    For i = 23 To 27: ws.Range("C" & i).Value = ws.Range(MACompColumnLetter & i).Value: Next i
    For i = 29 To 33: ws.Range("C" & i).Value = ws.Range(MACompColumnLetter & i).Value: Next i
    For i = 35 To 39: ws.Range("C" & i).Value = ws.Range(MACompColumnLetter & i).Value: Next i
    For i = 41 To 45: ws.Range("C" & i).Value = ws.Range(MACompColumnLetter & i).Value: Next i
    For i = 47 To 51: ws.Range("C" & i).Value = ws.Range(MACompColumnLetter & i).Value: Next i
    For i = 55 To 56: ws.Range("C" & i).Value = ws.Range(MACompColumnLetter & i).Value: Next i
    For i = 58 To 59: ws.Range("C" & i).Value = ws.Range(MACompColumnLetter & i).Value: Next i
    For i = 61 To 62: ws.Range("C" & i).Value = ws.Range(MACompColumnLetter & i).Value: Next i
    For i = 64 To 65: ws.Range("C" & i).Value = ws.Range(MACompColumnLetter & i).Value: Next i
    For i = 71 To 72: ws.Range("C" & i).Value = ws.Range(MACompColumnLetter & i).Value: Next i
    For i = 76 To 80: ws.Range("C" & i).Value = ws.Range(MACompColumnLetter & i).Value: Next i
    For i = 84 To 85: ws.Range("C" & i).Value = ws.Range(MACompColumnLetter & i).Value: Next i
    For i = 87 To 88: ws.Range("C" & i).Value = ws.Range(MACompColumnLetter & i).Value: Next i
    For i = 90 To 91: ws.Range("C" & i).Value = ws.Range(MACompColumnLetter & i).Value: Next i
    For i = 97 To 102: ws.Range("C" & i).Value = ws.Range(MACompColumnLetter & i).Value: Next i
    For i = 104 To 110: ws.Range("C" & i).Value = ws.Range(MACompColumnLetter & i).Value: Next i
    For i = 114 To 119: ws.Range("C" & i).Value = ws.Range(MACompColumnLetter & i).Value: Next i
    For i = 121 To 127: ws.Range("C" & i).Value = ws.Range(MACompColumnLetter & i).Value: Next i
    For i = 132 To 133: ws.Range("C" & i).Value = ws.Range(MACompColumnLetter & i).Value: Next i
    
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' Sub: MACompFinalResults3
' ----------------------------------------------------------------------------
' PURPOSE:
'   Transfers final computation results from the selected column to column G.
'   Used for secondary comparison or alternate determination.
'
' V1 EQUIVALENT: MA_Comp_finalresults3()
' ----------------------------------------------------------------------------
Public Sub MACompFinalResults3()
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim i As Long
    
    Set ws = ActiveSheet
    
    ' Copy main data
    ws.Range("G8").Value = ws.Range(MACompColumnLetter & "8").Value
    ws.Range("G144").Value = ws.Range(MACompColumnLetter & "144").Value
    
    ' Copy category and status code based on column
    Select Case MACompColumnLetter
        Case "D", "E", "F"
            ws.Range("G6").Value = ws.Range("E6").Value
            ws.Range("G7").Value = ws.Range("E7").Value
        Case "H", "I", "J"
            ws.Range("G6").Value = ws.Range("I6").Value
            ws.Range("G7").Value = ws.Range("I7").Value
        Case "K", "L", "M"
            ws.Range("G6").Value = ws.Range("L6").Value
            ws.Range("G7").Value = ws.Range("L7").Value
    End Select
    
    ' Copy data rows - grouped by section
    For i = 10 To 15: ws.Range("G" & i).Value = ws.Range(MACompColumnLetter & i).Value: Next i
    For i = 17 To 21: ws.Range("G" & i).Value = ws.Range(MACompColumnLetter & i).Value: Next i
    For i = 23 To 27: ws.Range("G" & i).Value = ws.Range(MACompColumnLetter & i).Value: Next i
    For i = 29 To 33: ws.Range("G" & i).Value = ws.Range(MACompColumnLetter & i).Value: Next i
    For i = 35 To 39: ws.Range("G" & i).Value = ws.Range(MACompColumnLetter & i).Value: Next i
    For i = 41 To 45: ws.Range("G" & i).Value = ws.Range(MACompColumnLetter & i).Value: Next i
    For i = 47 To 51: ws.Range("G" & i).Value = ws.Range(MACompColumnLetter & i).Value: Next i
    For i = 55 To 56: ws.Range("G" & i).Value = ws.Range(MACompColumnLetter & i).Value: Next i
    For i = 58 To 59: ws.Range("G" & i).Value = ws.Range(MACompColumnLetter & i).Value: Next i
    For i = 61 To 62: ws.Range("G" & i).Value = ws.Range(MACompColumnLetter & i).Value: Next i
    For i = 64 To 65: ws.Range("G" & i).Value = ws.Range(MACompColumnLetter & i).Value: Next i
    For i = 71 To 72: ws.Range("G" & i).Value = ws.Range(MACompColumnLetter & i).Value: Next i
    For i = 76 To 80: ws.Range("G" & i).Value = ws.Range(MACompColumnLetter & i).Value: Next i
    For i = 84 To 85: ws.Range("G" & i).Value = ws.Range(MACompColumnLetter & i).Value: Next i
    For i = 87 To 88: ws.Range("G" & i).Value = ws.Range(MACompColumnLetter & i).Value: Next i
    For i = 90 To 91: ws.Range("G" & i).Value = ws.Range(MACompColumnLetter & i).Value: Next i
    For i = 97 To 102: ws.Range("G" & i).Value = ws.Range(MACompColumnLetter & i).Value: Next i
    For i = 104 To 110: ws.Range("G" & i).Value = ws.Range(MACompColumnLetter & i).Value: Next i
    For i = 114 To 119: ws.Range("G" & i).Value = ws.Range(MACompColumnLetter & i).Value: Next i
    For i = 121 To 127: ws.Range("G" & i).Value = ws.Range(MACompColumnLetter & i).Value: Next i
    For i = 132 To 133: ws.Range("G" & i).Value = ws.Range(MACompColumnLetter & i).Value: Next i
    
    On Error GoTo 0
End Sub


' ============================================================================
' FORM DISPLAY FUNCTIONS
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: ShowMACompForm2
' ----------------------------------------------------------------------------
' PURPOSE:
'   Displays the MA Computation helper form 2.
'
' V1 EQUIVALENT: show_formMAC2()
' ----------------------------------------------------------------------------
Public Sub ShowMACompForm2()
    On Error Resume Next
    UF_MA_Comp2.Show
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' Sub: RedisplayMACompForm2
' ----------------------------------------------------------------------------
' PURPOSE:
'   Redisplays the MA Computation helper form 2.
'
' V1 EQUIVALENT: redisplayformMAC2()
' ----------------------------------------------------------------------------
Public Sub RedisplayMACompForm2()
    On Error Resume Next
    UF_MA_Comp2.Show
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' Sub: ShowMACompForm3
' ----------------------------------------------------------------------------
' PURPOSE:
'   Displays the MA Computation helper form 3.
'
' V1 EQUIVALENT: show_formMAC3()
' ----------------------------------------------------------------------------
Public Sub ShowMACompForm3()
    On Error Resume Next
    UF_MA_Comp3.Show
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' Sub: RedisplayMACompForm3
' ----------------------------------------------------------------------------
' PURPOSE:
'   Redisplays the MA Computation helper form 3.
'
' V1 EQUIVALENT: redisplayformMAC3()
' ----------------------------------------------------------------------------
Public Sub RedisplayMACompForm3()
    On Error Resume Next
    UF_MA_Comp3.Show
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' Sub: ShowMASelectForms
' ----------------------------------------------------------------------------
' PURPOSE:
'   Displays the MA form selection dialog.
' ----------------------------------------------------------------------------
Public Sub ShowMASelectForms()
    On Error Resume Next
    UF_MA_SelectForms.Show
    On Error GoTo 0
End Sub


' ============================================================================
' BUTTON CLEARING FUNCTIONS
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: ClearMAButtons
' ----------------------------------------------------------------------------
' PURPOSE:
'   Clears option buttons in a specified range. Used to reset form sections.
'
' PARAMETERS:
'   startNum - Starting button number
'   endNum   - Ending button number
'   targetCell - Cell to select after clearing (for navigation)
'
' V1 EQUIVALENT: MAClearButtons1to15, MAClearButtons16to24, etc.
' ----------------------------------------------------------------------------
Public Sub ClearMAButtons(ByVal startNum As Long, _
                          ByVal endNum As Long, _
                          ByVal targetCell As String)
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim i As Long
    
    Set ws = ActiveSheet
    
    For i = startNum To endNum
        ws.Shapes("OB " & i).Select
        With Selection
            .Value = xlOff
        End With
    Next i
    
    ws.Range(targetCell).Select
    
    On Error GoTo 0
End Sub

' Convenience wrappers for specific button ranges
Public Sub ClearMAButtons1to15()
    Call ClearMAButtons(1, 15, "AI102")
End Sub

Public Sub ClearMAButtons16to24()
    Call ClearMAButtons(16, 24, "AI115")
End Sub

Public Sub ClearMAButtons25to27()
    Call ClearMAButtons(25, 27, "AI305")
End Sub

Public Sub ClearMAButtons28to33()
    Call ClearMAButtons(28, 33, "AI422")
End Sub

Public Sub ClearMAButtons34to45()
    Call ClearMAButtons(34, 45, "AI521")
End Sub

Public Sub ClearMAButtons46to57()
    Call ClearMAButtons(46, 57, "AI643")
End Sub

Public Sub ClearMAButtons58to66()
    Call ClearMAButtons(58, 66, "AI754")
End Sub

Public Sub ClearMAButtons67to78()
    Call ClearMAButtons(67, 78, "AI851")
End Sub

Public Sub ClearMAButtons79to93()
    Call ClearMAButtons(79, 93, "AI967")
End Sub

Public Sub ClearMAButtons94to105()
    Call ClearMAButtons(94, 105, "AI1076")
End Sub

Public Sub ClearMAButtons106to117()
    Call ClearMAButtons(106, 117, "AI1192")
End Sub

Public Sub ClearMAButtons118to126()
    Call ClearMAButtons(118, 126, "AI1305")
End Sub

Public Sub ClearMAButtons127to132()
    Call ClearMAButtons(127, 132, "AI1408")
End Sub


' ============================================================================
' DATA TRANSFER FUNCTIONS
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: TransferMASectionA
' ----------------------------------------------------------------------------
' PURPOSE:
'   Transfers wage data from the wages section to Section A of the income comp.
'
' V1 EQUIVALENT: MA_Comp_Transfer_SectionA()
' ----------------------------------------------------------------------------
Public Sub TransferMASectionA()
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim linea As String, lineb As String, linec As String
    Dim ans As VbMsgBoxResult
    
    Set ws = ActiveSheet
    
    linea = ws.Range("A71").Value
    lineb = ws.Range("E71").Value
    linec = ws.Range("I71").Value
    
    ans = MsgBox("Any information previously entered into 'Section A' will be cleared " & _
                 "and updated with the new information from this section. " & _
                 "Are you sure you want to proceed?", vbYesNo)
    If ans = vbNo Then Exit Sub
    
    ' Clear Section A
    ws.Range("A15:I19").ClearContents
    
    ' First Member
    If ws.Range("A71").Value <> "LN #" Then
        ws.Range("A15").Select
        Do Until ActiveCell.Value = ""
            ActiveCell.Offset(1, 0).Select
        Loop
        ActiveCell.Value = Right(linea, 2) & "/311"
        ActiveCell.Offset(0, 2).Value = ws.Range("D81").Value
        ws.Range("D13").Value = ws.Range("A73").Value
        
        ' Additional periods if data exists
        If ws.Range("D90").Value <> 0 Then
            ActiveCell.Offset(0, 3).Value = ws.Range("D90").Value
            ws.Range("E13").Value = ws.Range("A82").Value
            ActiveCell.Offset(0, 4).Value = ws.Range("D99").Value
            ws.Range("F13").Value = ws.Range("A91").Value
        End If
        
        If ws.Range("D108").Value <> 0 Or ws.Range("D117").Value <> 0 Or ws.Range("D126").Value <> 0 Then
            ActiveCell.Offset(0, 5).Value = ws.Range("D108").Value
            ws.Range("G13").Value = ws.Range("A100").Value
            ActiveCell.Offset(0, 6).Value = ws.Range("D117").Value
            ws.Range("H13").Value = ws.Range("A109").Value
            ActiveCell.Offset(0, 7).Value = ws.Range("D126").Value
            ws.Range("I13").Value = ws.Range("A118").Value
        End If
    End If
    
    ' Second Member
    If ws.Range("E71").Value <> "LN #" Then
        ws.Range("A15").Select
        Do Until ActiveCell.Value = ""
            ActiveCell.Offset(1, 0).Select
        Loop
        ActiveCell.Value = Right(lineb, 2) & "/311"
        ActiveCell.Offset(0, 2).Value = ws.Range("H81").Value
        
        If ws.Range("H90").Value <> 0 Then
            ActiveCell.Offset(0, 3).Value = ws.Range("H90").Value
            ws.Range("E13").Value = ws.Range("E82").Value
            ActiveCell.Offset(0, 4).Value = ws.Range("H99").Value
            ws.Range("F13").Value = ws.Range("E91").Value
        End If
        
        If ws.Range("H108").Value <> 0 Or ws.Range("H117").Value <> 0 Or ws.Range("H126").Value <> 0 Then
            ActiveCell.Offset(0, 5).Value = ws.Range("H108").Value
            ws.Range("G13").Value = ws.Range("E100").Value
            ActiveCell.Offset(0, 6).Value = ws.Range("H117").Value
            ws.Range("H13").Value = ws.Range("E109").Value
            ActiveCell.Offset(0, 7).Value = ws.Range("H126").Value
            ws.Range("I13").Value = ws.Range("E118").Value
        End If
    End If
    
    ' Third Member
    If ws.Range("I71").Value <> "LN #" Then
        ws.Range("A15").Select
        Do Until ActiveCell.Value = ""
            ActiveCell.Offset(1, 0).Select
        Loop
        ActiveCell.Value = Right(linec, 2) & "/311"
        ActiveCell.Offset(0, 2).Value = ws.Range("L81").Value
        
        If ws.Range("L90").Value <> 0 Then
            ActiveCell.Offset(0, 3).Value = ws.Range("L90").Value
            ws.Range("E13").Value = ws.Range("I82").Value
            ActiveCell.Offset(0, 4).Value = ws.Range("L99").Value
            ws.Range("F13").Value = ws.Range("I91").Value
        End If
        
        If ws.Range("L108").Value <> 0 Or ws.Range("L117").Value <> 0 Or ws.Range("L126").Value <> 0 Then
            ActiveCell.Offset(0, 5).Value = ws.Range("L108").Value
            ws.Range("G13").Value = ws.Range("I100").Value
            ActiveCell.Offset(0, 6).Value = ws.Range("L117").Value
            ws.Range("H13").Value = ws.Range("I109").Value
            ActiveCell.Offset(0, 7).Value = ws.Range("L126").Value
            ws.Range("I13").Value = ws.Range("I118").Value
        End If
    End If
    
    ' Fill zeros where needed for admin period
    If ws.Range("E15").Value <> "" Or ws.Range("F15").Value <> "" Or _
       ws.Range("E16").Value <> "" Or ws.Range("F16").Value <> "" Or _
       ws.Range("E17").Value <> "" Or ws.Range("F17").Value <> "" Then
        If ws.Range("D15").Value <> "" Then
            ws.Range("E15").Value = ws.Range("D90").Value
            ws.Range("F15").Value = ws.Range("D99").Value
        End If
        If ws.Range("D16").Value <> "" Then
            ws.Range("E16").Value = ws.Range("H90").Value
            ws.Range("F16").Value = ws.Range("H99").Value
        End If
        If ws.Range("D17").Value <> "" Then
            ws.Range("E17").Value = ws.Range("L90").Value
            ws.Range("F17").Value = ws.Range("L99").Value
        End If
    End If
    
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' Sub: ClearMAWages
' ----------------------------------------------------------------------------
' PURPOSE:
'   Clears the MA wages worksheet section.
'
' V1 EQUIVALENT: MA_Comp_Clear_Wages()
' ----------------------------------------------------------------------------
Public Sub ClearMAWages()
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim ans As VbMsgBoxResult
    
    Set ws = ActiveSheet
    
    ans = MsgBox("Any information previously entered into this section will be deleted. " & _
                 "Are you sure you want to proceed?", vbYesNo)
    If ans = vbNo Then Exit Sub
    
    ' Clear wage data ranges
    ws.Range("A75:L79").ClearContents
    ws.Range("A84:L88").ClearContents
    ws.Range("A93:L97").ClearContents
    ws.Range("A102:L106").ClearContents
    ws.Range("A111:L115").ClearContents
    ws.Range("A120:L124").ClearContents
    
    ' Reset line number headers
    ws.Range("A71").Value = "LN #"
    ws.Range("E71").Value = "LN #"
    ws.Range("I71").Value = "LN #"
    
    ' Reset employer headers
    ws.Range("A72").Value = "Employer 1"
    ws.Range("C72").Value = "Employer 2"
    ws.Range("E72").Value = "Employer 1"
    ws.Range("G72").Value = "Employer 2"
    ws.Range("I72").Value = "Employer 1"
    ws.Range("K72").Value = "Employer 2"
    
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' Sub: ClearMAIncome
' ----------------------------------------------------------------------------
' PURPOSE:
'   Clears the MA Income Comp worksheet.
'
' V1 EQUIVALENT: MA_Comp_Clear_Income()
' ----------------------------------------------------------------------------
Public Sub ClearMAIncome()
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim ans As VbMsgBoxResult
    
    Set ws = ActiveSheet
    
    ans = MsgBox("Any information previously entered into this worksheet will be deleted. " & _
                 "Are you sure you want to proceed?", vbYesNo)
    If ans = vbNo Then Exit Sub
    
    ' Clear income data ranges
    ws.Range("A15:L19").ClearContents
    ws.Range("P25:V25").ClearContents
    ws.Range("P16:V16").ClearContents
    ws.Range("C24:L35").ClearContents
    ws.Range("C40:L45").ClearContents
    ws.Range("C52:L55").ClearContents
    ws.Range("C46:L46").ClearContents
    ws.Range("D13:L13").ClearContents
    
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' Sub: TransferMAToWorkbook
' ----------------------------------------------------------------------------
' PURPOSE:
'   Transfers income values from the MA Income Comp sheet to the MA Workbook.
'
' V1 EQUIVALENT: MA_Comp_Transfer_Workbook()
' ----------------------------------------------------------------------------
Public Sub TransferMAToWorkbook()
    On Error Resume Next
    
    Dim finalCol As String
    Dim maWorkbook As Worksheet
    Dim maIncomeComp As Worksheet
    
    Set maWorkbook = ThisWorkbook.Worksheets("MA Workbook")
    Set maIncomeComp = ThisWorkbook.Worksheets("MA Income Comp")
    
    finalCol = InputBox("Which column do you want to be transferred into Elements 371/372?" & _
                        vbNewLine & "Example: Q, R, or S?", "Final")
    If finalCol = "" Then Exit Sub
    
    finalCol = UCase(finalCol)
    
    If finalCol = "Q" Or finalCol = "R" Or finalCol = "S" Then
        maWorkbook.Range("AB1425").Value = maIncomeComp.Range(finalCol & "17").Value
        maWorkbook.Range("AB1430").Value = maIncomeComp.Range(finalCol & "19").Value
        maWorkbook.Range("AB1432").Value = maIncomeComp.Range(finalCol & "22").Value
        maWorkbook.Range("AB1434").Value = maIncomeComp.Range(finalCol & "24").Value
        maWorkbook.Shapes("Check Box 570").OLEFormat.Object.Value = 1
        maWorkbook.Shapes("Check Box 571").OLEFormat.Object.Value = 1
        maWorkbook.Shapes("Check Box 572").OLEFormat.Object.Value = 1
        maWorkbook.Shapes("Check Box 669").OLEFormat.Object.Value = 1
        maWorkbook.Shapes("Check Box 671").OLEFormat.Object.Value = 1
    Else
        MsgBox "You did not enter a Q or R or S. Please try again"
        Exit Sub
    End If
    
    ' QC Column
    If maWorkbook.Shapes("Check Box 565").OLEFormat.Object.Value <> 1 Then
        maWorkbook.Range("M1425").Value = maIncomeComp.Range("P17").Value
        maWorkbook.Range("M1428").Value = maIncomeComp.Range("P19").Value
        maWorkbook.Range("M1430").Value = maIncomeComp.Range("P22").Value
        maWorkbook.Range("M1432").Value = maIncomeComp.Range("P24").Value
        maWorkbook.Shapes("Check Box 566").OLEFormat.Object.Value = 1
        maWorkbook.Shapes("Check Box 567").OLEFormat.Object.Value = 1
        maWorkbook.Shapes("Check Box 568").OLEFormat.Object.Value = 1
        maWorkbook.Shapes("Check Box 670").OLEFormat.Object.Value = 1
        maWorkbook.Shapes("Check Box 672").OLEFormat.Object.Value = 1
    End If
    
    On Error GoTo 0
End Sub


' ============================================================================
' RESOURCES EXCLUSION FUNCTION
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: ToggleResourcesExclusion
' ----------------------------------------------------------------------------
' PURPOSE:
'   Toggles resource exclusion checkboxes based on CB 354 value.
'
' V1 EQUIVALENT: resources_ex()
' ----------------------------------------------------------------------------
Public Sub ToggleResourcesExclusion()
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim excludeValue As Long
    
    Set ws = ActiveSheet
    
    If ws.Shapes("CB 354").OLEFormat.Object.Value = 1 Then
        excludeValue = 1
    Else
        excludeValue = 0
    End If
    
    ' Toggle all related checkboxes
    ws.Shapes("CB 353").OLEFormat.Object.Value = excludeValue
    ws.Shapes("CB 700").OLEFormat.Object.Value = excludeValue
    ws.Shapes("CB 703").OLEFormat.Object.Value = excludeValue
    ws.Shapes("CB 708").OLEFormat.Object.Value = excludeValue
    ws.Shapes("CB 707").OLEFormat.Object.Value = excludeValue
    ws.Shapes("CB 729").OLEFormat.Object.Value = excludeValue
    ws.Shapes("CB 730").OLEFormat.Object.Value = excludeValue
    ws.Shapes("CB 537").OLEFormat.Object.Value = excludeValue
    ws.Shapes("CB 536").OLEFormat.Object.Value = excludeValue
    ws.Shapes("CB 3249").OLEFormat.Object.Value = excludeValue
    ws.Shapes("CB 3245").OLEFormat.Object.Value = excludeValue
    ws.Shapes("CB 3268").OLEFormat.Object.Value = excludeValue
    ws.Shapes("CB 273").OLEFormat.Object.Value = excludeValue
    ws.Shapes("CB 4519").OLEFormat.Object.Value = excludeValue
    ws.Shapes("CB 4520").OLEFormat.Object.Value = excludeValue
    
    If excludeValue = 1 Then
        ws.Range("G922").Select
    End If
    
    On Error GoTo 0
End Sub


' ============================================================================
' TPL (THIRD PARTY LIABILITY) FUNCTIONS
' ============================================================================

' ----------------------------------------------------------------------------
' Sub: FillTPLSection
' ----------------------------------------------------------------------------
' PURPOSE:
'   Fills the Third Party Liability section with data from the workbook.
'
' V1 EQUIVALENT: TPL_Fill()
' ----------------------------------------------------------------------------
Public Sub FillTPLSection()
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim maWorkbook As Worksheet
    Dim maIncomeComp As Worksheet
    Dim linea As String, lineb As String
    Dim wlinea As String, wlineb As String
    Dim wlineaname As String, wlinebname As String
    Dim i As Long
    
    Set ws = ActiveSheet
    Set maWorkbook = ThisWorkbook.Worksheets("MA Workbook")
    Set maIncomeComp = ThisWorkbook.Worksheets("MA Income Comp")
    
    ws.Range("G11").Value = maWorkbook.Range("F39").Value
    ws.Range("A13").Value = maWorkbook.Range("D20").Value
    
    ' Medicare Claim # from Workbook
    If maWorkbook.Range("I1092").Value <> "" Then
        ws.Range("F19").Value = maWorkbook.Range("I1092").Value
        ws.Range("C19").Value = ws.Range("A7").Value
        ws.Shapes("Check Box 33").OLEFormat.Object.Value = 1
        ws.Shapes("Check Box 31").OLEFormat.Object.Value = 1
    End If
    
    ' #2 Category
    ws.Shapes("Check Box 3").OLEFormat.Object.Value = 1
    ' Question 1 in Column A
    ws.Shapes("Check Box 43").OLEFormat.Object.Value = 1
    ' Question 2 in Column A
    ws.Shapes("Check Box 13").OLEFormat.Object.Value = 1
    ws.Shapes("Check Box 11").OLEFormat.Object.Value = 1
    ws.Shapes("Check Box 12").OLEFormat.Object.Value = 1
    
    ' Employer from wages section
    If maIncomeComp.Range("A72").Value <> "Employer 1" Then
        linea = maIncomeComp.Range("A71").Value
        wlinea = Right(linea, 2)
        For i = 11 To 20
            If maWorkbook.Cells(i, 10).Value = wlinea Then
                wlineaname = maWorkbook.Cells(i, 12).Value
            End If
        Next i
        ws.Range("C26").Value = wlineaname
        ws.Range("C29").Value = maIncomeComp.Range("A72").Value
    End If
    
    If maIncomeComp.Range("E72").Value <> "Employer 1" Then
        lineb = maIncomeComp.Range("E71").Value
        wlineb = Right(lineb, 2)
        For i = 11 To 20
            If maWorkbook.Cells(i, 10).Value = wlineb Then
                wlinebname = maWorkbook.Cells(i, 12).Value
            End If
        Next i
        ws.Range("G26").Value = wlinebname
        ws.Range("G29").Value = maIncomeComp.Range("E72").Value
    End If
    
    ' Question 3 in Column A
    ws.Shapes("Check Box 23").OLEFormat.Object.Value = 1
    ' Question 4 in Column A
    ws.Shapes("Check Box 15").OLEFormat.Object.Value = 1
    ws.Shapes("Check Box 17").OLEFormat.Object.Value = 1
    
    ' Question 5 - Absent parent from LRR section
    If maWorkbook.Range("L27").Value <> "" Then
        ws.Shapes("Check Box 18").OLEFormat.Object.Value = 1
        ws.Range("C45").Value = "5. Name: " & maWorkbook.Range("L27").Value
        ws.Range("C46").Value = "Address: " & maWorkbook.Range("AE27").Value & ", " & maWorkbook.Range("AE28").Value
        ws.Range("H45").Value = "SSN: " & maWorkbook.Range("AA27").Value
    End If
    
    ' Question 6 in Column A
    ws.Shapes("Check Box 21").OLEFormat.Object.Value = 1
    
    ws.Range("A19").Select
    
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' Sub: ClearTPLSection
' ----------------------------------------------------------------------------
' PURPOSE:
'   Clears the Third Party Liability section.
'
' V1 EQUIVALENT: TPL_Clear()
' ----------------------------------------------------------------------------
Public Sub ClearTPLSection()
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim ans As VbMsgBoxResult
    
    Set ws = ActiveSheet
    
    ans = MsgBox("Any information previously entered into the Third Party sheet " & _
                 "by the above button will be cleared. Are you sure you want to proceed?", vbYesNo)
    If ans = vbNo Then Exit Sub
    
    ws.Range("G11:I11").ClearContents
    ws.Range("A13:A14").ClearContents
    ws.Range("F19:I19").ClearContents
    ws.Range("C19:E19").ClearContents
    ws.Shapes("Check Box 33").OLEFormat.Object.Value = 0
    ws.Shapes("Check Box 31").OLEFormat.Object.Value = 0
    ws.Shapes("Check Box 3").OLEFormat.Object.Value = 0
    ws.Shapes("Check Box 43").OLEFormat.Object.Value = 0
    ws.Shapes("Check Box 13").OLEFormat.Object.Value = 0
    ws.Shapes("Check Box 11").OLEFormat.Object.Value = 0
    ws.Shapes("Check Box 12").OLEFormat.Object.Value = 0
    ws.Range("C26:F27").ClearContents
    ws.Range("C29:F30").ClearContents
    ws.Range("G26:I27").ClearContents
    ws.Range("G29:I30").ClearContents
    ws.Shapes("Check Box 23").OLEFormat.Object.Value = 0
    ws.Shapes("Check Box 15").OLEFormat.Object.Value = 0
    ws.Shapes("Check Box 17").OLEFormat.Object.Value = 0
    ws.Shapes("Check Box 18").OLEFormat.Object.Value = 0
    ws.Range("C45:G45").Value = "5. Name: "
    ws.Range("C46:G46").Value = "Address: "
    ws.Range("H45:I45").Value = "SSN: "
    ws.Shapes("Check Box 21").OLEFormat.Object.Value = 0
    
    ws.Range("A19").Select
    
    On Error GoTo 0
End Sub


