VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_PopulateMain 
   Caption         =   "Populate Review Schedule"
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   4680
   OleObjectBlob   =   "UF_PopulateMain.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_PopulateMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ============================================================================
' UF_PopulateMain - Main Population Form
' ============================================================================
' WHAT THIS FORM DOES:
'   This is the main entry point for population. The user selects:
'   1. Program type (SNAP Positive, SNAP Negative, TANF, MA, GA)
'   2. Review month (YYYYMM format from dropdown)
'   3. Output type (Schedule or Transmittal)
'
'   After clicking OK, it calls either Review_Schedule or Transmittal
'   based on the user's selection.
'
' FORM CONTROLS:
'   - ComboBox1: Program selection dropdown
'   - ComboBox2: Month selection dropdown
'   - OptionButton1: "Schedule" radio button
'   - OptionButton2: "Transmittal" radio button
'   - CommandButton1: OK button
'   - CommandButton2: Cancel button
'
' HOW TO USE:
'   1. User runs macro that shows this form
'   2. User selects program and month from dropdowns
'   3. User selects Schedule or Transmittal option
'   4. User clicks OK to start population
'
' VALIDATION:
'   - Both dropdowns must have values selected
'   - One of the radio buttons must be selected
'   - If validation fails, form is re-displayed
'
' DEPENDENCIES:
'   - Pop_Main.Review_Schedule: Creates individual review schedules
'   - Pop_Transmittals.Transmittal: Creates batch transmittal sheets
'   - Populate sheet: Stores selected values in cells W7 and X7
'
' CHANGE LOG:
'   2026-01-03  Created from UserForm50.frm
'               Added V2 comments, renamed to UF_PopulateMain
' ============================================================================

Option Explicit

' ----------------------------------------------------------------------------
' Event: CommandButton1_Click (OK Button)
' ----------------------------------------------------------------------------
' PURPOSE:
'   Validates user input and initiates the selected operation.
' ----------------------------------------------------------------------------
Private Sub CommandButton1_Click()
    ' Validate that both dropdowns have selections
    If UF_PopulateMain.ComboBox1.Value = "" Or UF_PopulateMain.ComboBox2.Value = "" Then
        MsgBox "Please fill in both Program and Month fields", vbExclamation
        UF_PopulateMain.Hide
        Call redisplayform50
        Exit Sub
    End If
    
    ' Validate that a radio button is selected
    If OptionButton1.Value = False And OptionButton2.Value = False Then
        MsgBox "Please select either Schedule or Transmittal", vbExclamation
        UF_PopulateMain.Hide
        Call redisplayform50
        Exit Sub
    End If
    
    ' Store selections in Populate sheet for use by population modules
    Worksheets("Populate").Range("W7").Value = UF_PopulateMain.ComboBox1.Value
    Worksheets("Populate").Range("X7").Value = UF_PopulateMain.ComboBox2.Value
    
    UF_PopulateMain.Hide
    
    ' Call the appropriate routine based on selection
    If OptionButton1.Value = True Then
        ' Schedule selected - create individual review schedules
        Call Review_Schedule
    ElseIf OptionButton2.Value = True Then
        ' Transmittal selected - create batch transmittal sheets
        Call Transmittal
    End If
End Sub

' ----------------------------------------------------------------------------
' Event: CommandButton2_Click (Cancel Button)
' ----------------------------------------------------------------------------
' PURPOSE:
'   Closes the form and ends the macro.
' ----------------------------------------------------------------------------
Private Sub CommandButton2_Click()
    UF_PopulateMain.Hide
    End
End Sub

' ----------------------------------------------------------------------------
' Helper: redisplayform50
' ----------------------------------------------------------------------------
' PURPOSE:
'   Re-displays this form after a validation error.
' ----------------------------------------------------------------------------
Private Sub redisplayform50()
    UF_PopulateMain.Show
End Sub


