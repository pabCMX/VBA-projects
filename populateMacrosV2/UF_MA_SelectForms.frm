VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_MA_SelectForms 
   Caption         =   "Select Form"
   ClientHeight    =   3225
   ClientLeft      =   80
   ClientTop       =   440
   ClientWidth     =   7210
   OleObjectBlob   =   "UF_MA_SelectForms.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_MA_SelectForms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ============================================================================
' UF_MA_SelectForms - MA-Specific Form Selection Dialog
' ============================================================================
' WHAT THIS FORM DOES:
'   Provides a simplified form selection dialog specifically for MA
'   (Medical Assistance) reviews. This is used in contexts where only
'   a subset of MA-specific forms are needed.
'
' HOW IT WORKS:
'   1. Form detects whether it's an MA Positive or MA Negative review
'   2. ListBox1 is populated with relevant forms
'   3. User selects a form and clicks OK
'   4. The appropriate generation function is called
'
' PROGRAM DETECTION:
'   Uses the first character of the sheet name (review number):
'   - "2" = MA Positive (20xxx, 21xxx)
'   - "8" = MA Negative (80xxx, 81xxx, 82xxx, 83xxx)
'
' FORM CONTROLS:
'   - ListBox1: List of available forms for the MA program type
'   - CommandButton1: OK - Generate selected form
'   - CommandButton2: Cancel - Close without action
'
' CALLED FUNCTIONS:
'   - MailMergeandSave: Generates findings memo via mail merge
'   - Taxonomy: Generates taxonomy information memo
'   - ComSpouse: Community spouse questionnaire
'   - QC14C: QC-14C form
'   - QC15: QC-15 form
'
' WHY SEPARATE FROM MAIN SELECT FORMS:
'   This form is used in specific MA workflow contexts where a smaller
'   subset of forms is appropriate. The main UF_SelectForms provides
'   the full menu of all available forms.
'
' CHANGE LOG:
'   2026-01-03  Renamed from MASelectForms to UF_MA_SelectForms
'               Added V2-style documentation
'               Preserved original V1 logic for compatibility
' ============================================================================

Option Explicit

' ----------------------------------------------------------------------------
' Event: ListBox1_Click
' ----------------------------------------------------------------------------
Private Sub ListBox1_Click()
    ' Event handler (currently unused)
End Sub

' ----------------------------------------------------------------------------
' Event: UserForm_Click
' ----------------------------------------------------------------------------
Private Sub UserForm_Click()
    ' Event handler (currently unused)
End Sub

' ----------------------------------------------------------------------------
' Event: CommandButton1_Click (OK Button)
' ----------------------------------------------------------------------------
' PURPOSE:
'   Handles the OK button click. Routes to the appropriate form generation
'   function based on the selected index.
' ----------------------------------------------------------------------------
Private Sub CommandButton1_Click()
    Dim i As Integer
    
    ' Find selected form in the list
    For i = 0 To ListBox1.ListCount - 1
        If ListBox1.Selected(i) = True Then
            Unload UF_MA_SelectForms
            
            ' Route to appropriate function based on index
            If i = 0 Then
               Call MailMergeandSave
            ElseIf i = 1 Then
                Call Taxonomy
            'ElseIf i = 2 Then
            '    Call MAAppt
            ElseIf i = 2 Then
                 Call ComSpouse
            ElseIf i = 3 Then
                 Call QC14C
            ElseIf i = 4 Then
                 Call QC15
            End If
        End If
    Next i
End Sub

' ----------------------------------------------------------------------------
' Event: CommandButton2_Click (Cancel Button)
' ----------------------------------------------------------------------------
Private Sub CommandButton2_Click()
    Unload UF_MA_SelectForms
    End
End Sub

' ----------------------------------------------------------------------------
' Event: UserForm_Initialize
' ----------------------------------------------------------------------------
' PURPOSE:
'   Populates the ListBox with forms available for the current MA program type.
' ----------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    Select Case Left(ActiveSheet.Name, 1)
        ' ==================================================================
        ' MA NEGATIVE (80xxx, 81xxx, 82xxx, 83xxx)
        ' ==================================================================
        Case "8"
            ListBox1.AddItem ("Findings Memo")
            ListBox1.AddItem ("Taxonomy Information Memo")
            
        ' ==================================================================
        ' MA POSITIVE (20xxx, 21xxx)
        ' ==================================================================
        Case "2"
            ListBox1.AddItem ("Findings Memo")
            ListBox1.AddItem ("Taxonomy Information Memo")
            'ListBox1.AddItem ("MA Appointment Letter")
            ListBox1.AddItem ("Community Spouse")
            ListBox1.AddItem ("QC 14")
            ListBox1.AddItem ("QC 15")
    End Select
End Sub

