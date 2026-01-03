VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_DatePicker 
   Caption         =   "Enter date for CAO Appointment"
   ClientHeight    =   3225
   ClientLeft      =   80
   ClientTop       =   440
   ClientWidth     =   5880
   OleObjectBlob   =   "UF_DatePicker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_DatePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ============================================================================
' UF_DatePicker - Date Selection Dialog
' ============================================================================
' WHAT THIS FORM DOES:
'   Provides a simple interface for selecting a date. Used primarily for
'   appointment letters where the examiner needs to choose an interview date.
'
' HOW IT WORKS:
'   1. User enters a date in TextBox1 in MM/DD/YYYY format
'   2. Form validates that the date is not in the past or on a weekend
'   3. Sets the module-level variable ApptDate with formatted date
'
' CONTROLS:
'   - TextBox1: Single textbox for date entry (MM/DD/YYYY format)
'   - CommandButton1: OK - Accept the date
'   - CommandButton2: Cancel - Close without selecting
'
' VALIDATION:
'   - Checks that date is not in the past
'   - Checks that date is not on a weekend (Saturday/Sunday)
'   - Shows error message and keeps form open if validation fails
'
' OUTPUT:
'   Sets the module-level variable ApptDate to formatted string:
'   Example: "Monday, January 15, 2024"
'
' WHY NOT USE A CALENDAR CONTROL:
'   The Microsoft Date and Time Picker control is not always available on
'   all systems (especially 64-bit Office). This simple textbox approach
'   ensures compatibility across all Office versions.
'
' CHANGE LOG:
'   2026-01-03  Renamed from SelectDate to UF_DatePicker
'               Added V2-style documentation
'               Preserved original V1 logic for compatibility
' ============================================================================

Option Explicit

' ----------------------------------------------------------------------------
' Event: CommandButton1_Click (OK Button)
' ----------------------------------------------------------------------------
' PURPOSE:
'   Validates the entered date and sets the ApptDate variable.
' ----------------------------------------------------------------------------
Private Sub CommandButton1_Click()
    Dim date_array As Variant
    Dim date_entered As Date
    
    ' Parse the date from MM/DD/YYYY format
    date_array = Split(TextBox1.Value, "/")
    date_entered = DateSerial(date_array(2), date_array(0), date_array(1))
    
    ' Validate: Check if date is in the past
    If date_entered < Date Then
        MsgBox "Appointment date is in the past. Please choose another date."
        ' Keep form open for correction
        Exit Sub
    End If
    
    ' Validate: Check if date is on a weekend
    If Weekday(date_entered) = 1 Or Weekday(date_entered) = 7 Then
        MsgBox "Appointment date is on a weekend. Please choose another date."
        ' Keep form open for correction
        Exit Sub
    End If
    
    ' Date is valid - format it and set the module variable
    ApptDate = Format(date_entered, "dddd, mmmm dd, yyyy")
    
    ' Close the form
    Unload Me
End Sub

' ----------------------------------------------------------------------------
' Event: CommandButton2_Click (Cancel Button)
' ----------------------------------------------------------------------------
' PURPOSE:
'   Closes the form without selecting a date.
' ----------------------------------------------------------------------------
Private Sub CommandButton2_Click()
    Unload Me
End Sub

' ----------------------------------------------------------------------------
' Event: UserForm_Activate
' ----------------------------------------------------------------------------
' PURPOSE:
'   Sets the default value to today's date when the form appears.
' ----------------------------------------------------------------------------
Private Sub UserForm_Activate()
    TextBox1.Value = Date
End Sub

