VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_GA_Helper2 
   Caption         =   "Final Determination"
   ClientHeight    =   3225
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   4710
   OleObjectBlob   =   "UF_GA_Helper2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_GA_Helper2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ============================================================================
' UF_GA_Helper2 - GA Final Determination Column Selector
' ============================================================================
' WHAT THIS FORM DOES:
'   Second helper form for GA (General Assistance) computations. Similar to
'   UF_GA_Helper1 but used for the final determination step.
'
' HOW IT WORKS:
'   1. User selects a column from ComboBox1
'   2. Selection is stored in GA Computation sheet cell AL78
'   3. Calls GAfinaldetermination to process the final values
'
' FORM CONTROLS:
'   - ComboBox1: Column selection dropdown
'   - CommandButton1: OK - Accept the selection
'   - CommandButton2: Cancel - Close without action
'
' OUTPUT:
'   - Stores selection in Worksheets("GA Computation").Range("AL78")
'   - Calls GAfinaldetermination to process the transfer
'
' DIFFERENCE FROM UF_GA_Helper1:
'   - Helper1: Used during initial calculation phase (AL77)
'   - Helper2: Used for final determination (AL78)
'
' CHANGE LOG:
'   2026-01-03  Renamed from GAUserForm2 to UF_GA_Helper2
'               Added V2-style documentation
'               Preserved original V1 logic for compatibility
' ============================================================================

Option Explicit

' ----------------------------------------------------------------------------
' Event: CommandButton1_Click (OK Button)
' ----------------------------------------------------------------------------
' PURPOSE:
'   Validates selection and calls the final determination function.
' ----------------------------------------------------------------------------
Private Sub CommandButton1_Click()
    ' Validate that a column is selected
    If UF_GA_Helper2.ComboBox1.Value = "" Then
        MsgBox "Please pick a column to place your results."
        UF_GA_Helper2.Hide
        Call redisplayGAform2
        Exit Sub
    End If
    
    ' Store selection in GA Computation sheet (different cell than Helper1)
    Worksheets("GA Computation").Range("AL78") = UF_GA_Helper2.ComboBox1.Value
    
    UF_GA_Helper2.Hide
    
    ' Call the final determination function
    ' V2: Call the renamed function in Review_GA_Elements
    Call Review_GA_Elements.GAfinaldetermination
End Sub

' ----------------------------------------------------------------------------
' Event: CommandButton2_Click (Cancel Button)
' ----------------------------------------------------------------------------
' PURPOSE:
'   Closes the form without processing.
' ----------------------------------------------------------------------------
Private Sub CommandButton2_Click()
    UF_GA_Helper2.Hide
    End
End Sub

' ----------------------------------------------------------------------------
' Event: UserForm_Click
' ----------------------------------------------------------------------------
Private Sub UserForm_Click()
    ' Event handler (currently unused)
End Sub

