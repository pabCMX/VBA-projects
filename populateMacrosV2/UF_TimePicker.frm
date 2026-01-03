VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_TimePicker 
   Caption         =   "UserForm3"
   ClientHeight    =   3225
   ClientLeft      =   80
   ClientTop       =   440
   ClientWidth     =   5880
   OleObjectBlob   =   "UF_TimePicker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_TimePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ============================================================================
' UF_TimePicker - Time Selection Dialog
' ============================================================================
' WHAT THIS FORM DOES:
'   Provides a simple interface for selecting an appointment time. Uses
'   a listbox with pre-populated times in 15-minute increments to minimize
'   input errors.
'
' HOW IT WORKS:
'   1. Form initializes with times from 8:00 AM to 4:00 PM in 15-min intervals
'   2. User selects a time from ListBox1
'   3. Sets the module-level variable ApptTime with selected value
'
' CONTROLS:
'   - ListBox1: List of available appointment times
'   - CommandButton1: OK - Accept the selected time
'   - CommandButton2: Cancel - Close without selecting
'
' TIME RANGE:
'   8:00 AM to 4:00 PM in 15-minute increments
'   (Typical CAO office hours)
'
' OUTPUT:
'   Sets the module-level variable ApptTime to the selected time string
'   Example: "10:00 AM"
'
' CHANGE LOG:
'   2026-01-03  Renamed from SelectTime to UF_TimePicker
'               Added V2-style documentation
'               Preserved original V1 logic for compatibility
' ============================================================================

Option Explicit

' ----------------------------------------------------------------------------
' Event: CommandButton1_Click (OK Button)
' ----------------------------------------------------------------------------
' PURPOSE:
'   Sets the ApptTime variable to the selected time.
' ----------------------------------------------------------------------------
Private Sub CommandButton1_Click()
    Dim i As Integer
    
    ' Find the selected time in the list
    For i = 0 To ListBox1.ListCount - 1
        If ListBox1.Selected(i) = True Then
            ApptTime = ListBox1.Value
            Unload UF_TimePicker
        End If
    Next i
End Sub

' ----------------------------------------------------------------------------
' Event: CommandButton2_Click (Cancel Button)
' ----------------------------------------------------------------------------
' PURPOSE:
'   Closes the form without selecting a time.
' ----------------------------------------------------------------------------
Private Sub CommandButton2_Click()
    Unload UF_TimePicker
    End
End Sub

' ----------------------------------------------------------------------------
' Event: UserForm_Initialize
' ----------------------------------------------------------------------------
' PURPOSE:
'   Populates the listbox with available appointment times.
' ----------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    ' Initialize times from 8:00 AM to 4:00 PM in 15-minute increments
    ListBox1.AddItem ("8:00 AM")
    ListBox1.AddItem ("8:15 AM")
    ListBox1.AddItem ("8:30 AM")
    ListBox1.AddItem ("8:45 AM")
    ListBox1.AddItem ("9:00 AM")
    ListBox1.AddItem ("9:15 AM")
    ListBox1.AddItem ("9:30 AM")
    ListBox1.AddItem ("9:45 AM")
    ListBox1.AddItem ("10:00 AM")
    ListBox1.AddItem ("10:15 AM")
    ListBox1.AddItem ("10:30 AM")
    ListBox1.AddItem ("10:45 AM")
    ListBox1.AddItem ("11:00 AM")
    ListBox1.AddItem ("11:15 AM")
    ListBox1.AddItem ("11:30 AM")
    ListBox1.AddItem ("11:45 AM")
    ListBox1.AddItem ("12:00 PM")
    ListBox1.AddItem ("12:15 PM")
    ListBox1.AddItem ("12:30 PM")
    ListBox1.AddItem ("12:45 PM")
    ListBox1.AddItem ("1:00 PM")
    ListBox1.AddItem ("1:15 pM")
    ListBox1.AddItem ("1:30 PM")
    ListBox1.AddItem ("1:45 PM")
    ListBox1.AddItem ("2:00 PM")
    ListBox1.AddItem ("2:15 pM")
    ListBox1.AddItem ("2:30 PM")
    ListBox1.AddItem ("2:45 PM")
    ListBox1.AddItem ("3:00 PM")
    ListBox1.AddItem ("3:15 pM")
    ListBox1.AddItem ("3:30 PM")
    ListBox1.AddItem ("3:45 PM")
    ListBox1.AddItem ("4:00 PM")
End Sub

