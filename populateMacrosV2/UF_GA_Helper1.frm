VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_GA_Helper1 
   Caption         =   "Where do Results Go?"
   ClientHeight    =   3210
   ClientLeft      =   50
   ClientTop       =   350
   ClientWidth     =   4710
   OleObjectBlob   =   "UF_GA_Helper1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_GA_Helper1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ============================================================================
' UF_GA_Helper1 - GA Computation Results Column Selector
' ============================================================================
' WHAT THIS FORM DOES:
'   First helper form for GA (General Assistance) computations. Allows the
'   examiner to select which column from the GA Computation sheet contains
'   the final results to be transferred to the main schedule.
'
' HOW IT WORKS:
'   1. User selects a column from ComboBox1
'   2. Selection is stored in GA Computation sheet cell AL77
'   3. Calls GAfinalresults to copy values from selected column
'
' FORM CONTROLS:
'   - ComboBox1: Column selection dropdown
'   - CommandButton1: OK - Accept the selection
'   - CommandButton2: Cancel - Close without action
'
' OUTPUT:
'   - Stores selection in Worksheets("GA Computation").Range("AL77")
'   - Calls GAfinalresults to process the transfer
'
' WHY COLUMN SELECTION:
'   GA computations may involve multiple scenarios (e.g., with/without
'   certain deductions). The examiner calculates all scenarios and then
'   selects which one represents the correct determination.
'
' CHANGE LOG:
'   2026-01-03  Renamed from GAUserForm1 to UF_GA_Helper1
'               Added V2-style documentation
'               Preserved original V1 logic for compatibility
' ============================================================================

Option Explicit

' ----------------------------------------------------------------------------
' Event: CommandButton1_Click (OK Button)
' ----------------------------------------------------------------------------
' PURPOSE:
'   Validates selection and calls the results transfer function.
' ----------------------------------------------------------------------------
Private Sub CommandButton1_Click()
    ' Validate that a column is selected
    If UF_GA_Helper1.ComboBox1.Value = "" Then
        MsgBox "Please pick a column to place your results."
        UF_GA_Helper1.Hide
        Call redisplayGAform1
        Exit Sub
    End If
    
    ' Store selection in GA Computation sheet
    Worksheets("GA Computation").Range("AL77") = UF_GA_Helper1.ComboBox1.Value
    
    UF_GA_Helper1.Hide
    
    ' Call the results transfer function
    ' V2: Call the renamed function in Review_GA_Elements
    Call Review_GA_Elements.GAfinalresults
End Sub

' ----------------------------------------------------------------------------
' Event: CommandButton2_Click (Cancel Button)
' ----------------------------------------------------------------------------
' PURPOSE:
'   Closes the form without processing.
' ----------------------------------------------------------------------------
Private Sub CommandButton2_Click()
    UF_GA_Helper1.Hide
    End
End Sub

' ----------------------------------------------------------------------------
' Event: UserForm_Click
' ----------------------------------------------------------------------------
Private Sub UserForm_Click()
    ' Event handler (currently unused)
End Sub

