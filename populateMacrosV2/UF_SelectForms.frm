VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_SelectForms 
   Caption         =   "Select Form"
   ClientHeight    =   3192
   ClientLeft      =   40
   ClientTop       =   180
   ClientWidth     =   7520
   OleObjectBlob   =   "UF_SelectForms.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_SelectForms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ============================================================================
' UF_SelectForms - Form and Memo Selection Dialog
' ============================================================================
' WHAT THIS FORM DOES:
'   This is the main form selection dialog that allows examiners to generate
'   various forms, memos, and letters needed for QC reviews. The available
'   options are filtered based on the program type of the active review.
'
' HOW IT WORKS:
'   1. Form detects the program type from the active sheet name
'   2. ListBox1 is populated with forms available for that program
'   3. User selects a form and clicks OK
'   4. The appropriate generation function is called based on the index
'
' PROGRAM DETECTION:
'   Uses the first 1-2 characters of the sheet name (review number):
'   - "5" = SNAP Positive (50xxx, 51xxx, 55xxx)
'   - "6" = SNAP Negative (60xxx, 61xxx, 65xxx, 66xxx)
'   - "1" = TANF (14xxx)
'   - "9" = GA (9xxxx)
'   - "2" = MA Positive (20xxx, 21xxx, 24xxx for PE)
'   - "8" = MA Negative (80xxx, 81xxx, 82xxx, 83xxx)
'
' FORM CONTROLS:
'   - ListBox1: List of available forms for the current program
'   - CommandButton1: OK - Generate selected form
'   - CommandButton2: Cancel - Close without action
'
' CALLED FUNCTIONS (by program):
'   The form calls various generation functions defined in other modules:
'   - Finding_Memo_sub: Findings memo (all programs)
'   - CAOAppt: CAO appointment letter (SNAP/MA)
'   - Rush: Case summary template
'   - Info: Information memo (Review_CashMemos module)
'   - And many more (see code comments for full list)
'
' CHANGE LOG:
'   2026-01-03  Renamed from SelectForms to UF_SelectForms
'               Added V2-style documentation
'               Preserved original V1 logic for compatibility
' ============================================================================

Option Explicit

' ----------------------------------------------------------------------------
' Event: CommandButton1_Click (OK Button)
' ----------------------------------------------------------------------------
' PURPOSE:
'   Handles the OK button click. Uses the V2 FormRouter to route to the
'   appropriate form generation function based on the selected form name.
'
' V2 CHANGE:
'   Now uses Review_FormRouter.GenerateForm() for centralized routing instead
'   of calling individual functions directly. This makes it easier to add new
'   forms and rename functions without modifying this UserForm.
' ----------------------------------------------------------------------------
Private Sub CommandButton1_Click()
    Dim i As Integer
    Dim selectedForm As String
    
    ' Find the selected form name from the listbox
    selectedForm = ""
    For i = 0 To ListBox1.ListCount - 1
        If ListBox1.Selected(i) = True Then
            selectedForm = ListBox1.List(i)
            Exit For
        End If
    Next i
    
    ' If nothing selected, exit
    If selectedForm = "" Then
        MsgBox "Please select a form to generate.", vbExclamation, "No Selection"
        Exit Sub
    End If
    
    ' Unload the form before calling the generation function
    Unload UF_SelectForms
    
    ' Use the centralized FormRouter to route to the correct function
    ' The FormRouter maps form names to their generation functions
    Call Review_FormRouter.GenerateForm(selectedForm)
End Sub

' ----------------------------------------------------------------------------
' Event: CommandButton2_Click (Cancel Button)
' ----------------------------------------------------------------------------
Private Sub CommandButton2_Click()
    Unload UF_SelectForms
    End
End Sub

' ----------------------------------------------------------------------------
' Event: ListBox1_Click
' ----------------------------------------------------------------------------
Private Sub ListBox1_Click()
    ' Event handler for list box selection (currently unused)
End Sub

' ----------------------------------------------------------------------------
' Event: UserForm_Initialize
' ----------------------------------------------------------------------------
' PURPOSE:
'   Runs when the form loads. Populates the ListBox with forms available
'   for the current program type.
' ----------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    Select Case Left(ActiveSheet.Name, 1)
        ' ==================================================================
        ' SNAP POSITIVE (50xxx, 51xxx, 55xxx)
        ' ==================================================================
        Case "5"
            ListBox1.AddItem ("Findings Memo")
            ListBox1.AddItem ("CAO Appointment Letter")
            ListBox1.AddItem ("Case Summary Template")
            ListBox1.AddItem ("Telephone Appointment Letter")
            ListBox1.AddItem ("SNAP QC14F Memo")
            ListBox1.AddItem ("SNAP QC14R Memo")
            ListBox1.AddItem ("SNAP QC14C Memo")
            ListBox1.AddItem ("Timeliness Information/Finding Memo")
            ListBox1.AddItem ("Information Memo")
            ListBox1.AddItem ("Post Office Memo")
            ListBox1.AddItem ("Pending Letter")
            ListBox1.AddItem ("Drop Worksheet")
            ListBox1.AddItem ("CAO Forms Request")
            ListBox1.AddItem ("Error Under Threshold Memo")
            ListBox1.AddItem ("LEP")
            ListBox1.AddItem ("Spanish CAO Appointment Letter")
            ListBox1.AddItem ("Spanish Telephone Appointment Letter")
            ListBox1.AddItem ("PA78")
            ListBox1.AddItem ("Spanish Pending Letter")
            
        ' ==================================================================
        ' SNAP NEGATIVE (60xxx, 61xxx, 65xxx, 66xxx)
        ' ==================================================================
        Case "6"
            ListBox1.AddItem ("Findings Memo")
            ListBox1.AddItem ("Case Summary Template")
            ListBox1.AddItem ("Negative Info Memo")
            ListBox1.AddItem ("LEP")
            
        ' ==================================================================
        ' MA NEGATIVE (80xxx, 81xxx, 82xxx, 83xxx)
        ' ==================================================================
        Case "8"
            ListBox1.AddItem ("Findings Memo")
            ListBox1.AddItem ("Deficiency Memo")
            ListBox1.AddItem ("Potential Error Call Memo")
            ListBox1.AddItem ("Information Memo")
            ListBox1.AddItem ("LEP")
            
        ' ==================================================================
        ' MA POSITIVE (20xxx, 21xxx) and MA PE (24xxx)
        ' ==================================================================
        Case "2"
            ' Special handling for PE cases
            If Left(ActiveSheet.Name, 2) = "24" Then
                ListBox1.AddItem ("Findings Memo")
                ListBox1.AddItem ("Information Memo")
                'ListBox1.AddItem ("PE Deficiency Letter")
            Else
                ListBox1.AddItem ("Findings Memo")
                ListBox1.AddItem ("Deficiency Memo")
                ListBox1.AddItem ("MA Appointment Letter")
                ListBox1.AddItem ("Community Spouse Questionaire")
                ListBox1.AddItem ("Potential Error Call Memo")
                ListBox1.AddItem ("Information Memo")
                ListBox1.AddItem ("QC14 Coop Memo")
                ListBox1.AddItem ("QC14 CAO Request Memo")
                ListBox1.AddItem ("QC15 Memo")
                ListBox1.AddItem ("Preliminary Information Memo")
                ListBox1.AddItem ("NH LRR")
                ListBox1.AddItem ("NH Home Business")
                ListBox1.AddItem ("PA472-Pending Letter")
                ListBox1.AddItem ("PA472-Self Emp.")
                ListBox1.AddItem ("Self Emp.")
                ListBox1.AddItem ("PA76")
                ListBox1.AddItem ("PA78")
                ListBox1.AddItem ("PA83-Z")
                ListBox1.AddItem ("Household Composition")
                ListBox1.AddItem ("SAVE Deficiency Memo")
                ListBox1.AddItem ("MA Support Form")
                ListBox1.AddItem ("MA Waiver Memo")
                ListBox1.AddItem ("Adoption Foster Care")
                ListBox1.AddItem ("Zero Income Request")
                ListBox1.AddItem ("QC14 LTC Waiver Memo")
                ListBox1.AddItem ("Funeral Home Letter")
                ListBox1.AddItem ("LEP")
                ListBox1.AddItem ("Spanish MA App. Letter")
            End If
            
        ' ==================================================================
        ' TANF (14xxx)
        ' ==================================================================
        Case "1"
            ListBox1.AddItem ("Findings Memo")
            ListBox1.AddItem ("Information Memo")
            ListBox1.AddItem ("Potential Error Call Memo")
            'ListBox1.AddItem ("Notification Signature Requirement Info Memo")
            ListBox1.AddItem ("Notification Requirement Info Memo")
            ListBox1.AddItem ("AMR")
            ListBox1.AddItem ("Criminal History")
            ListBox1.AddItem ("TANF CAO Request Form")
            ListBox1.AddItem ("SAVE")
            ListBox1.AddItem ("Employment/Earnings Report")
            ListBox1.AddItem ("LEP")
            ListBox1.AddItem ("TANF Pending")
            ListBox1.AddItem ("School Verification")
            
        ' ==================================================================
        ' GA (9xxxx)
        ' ==================================================================
        Case "9"
            ListBox1.AddItem ("Findings Memo")
            ListBox1.AddItem ("AMR Memo")
            ListBox1.AddItem ("Criminal History Memo")
            ListBox1.AddItem ("Information Memo")
            ListBox1.AddItem ("Potential Error Call Memo")
           ' ListBox1.AddItem ("GA CAO Request Form")
            ListBox1.AddItem ("SAVE")
            ListBox1.AddItem ("Employment/Earnings Report")
    End Select
End Sub

