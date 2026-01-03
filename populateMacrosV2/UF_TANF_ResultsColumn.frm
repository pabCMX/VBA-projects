VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_TANF_ResultsColumn 
   Caption         =   "Where do Results Go?"
   ClientHeight    =   3210
   ClientLeft      =   80
   ClientTop       =   450
   ClientWidth     =   5880
   OleObjectBlob   =   "UF_TANF_ResultsColumn.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_TANF_ResultsColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ============================================================================
' UF_TANF_ResultsColumn - TANF Computation Results Column Selector
' ============================================================================
' WHAT THIS FORM DOES:
'   After completing the TANF Computation sheet, the examiner needs to
'   choose which column's results to transfer to the main schedule.
'   
'   The TANF Computation sheet has two scenarios:
'   - Column D: First calculation scenario
'   - Column E: Second/comparison scenario
'
'   This form asks the examiner which column contains the correct results
'   to be copied to the final schedule.
'
' FORM CONTROLS:
'   - ComboBox1: Column selection (D or E)
'   - CommandButton1: OK - proceeds with result transfer
'   - CommandButton2: Cancel - aborts the operation
'
' WHY TWO COLUMNS?
'   Examiners often need to calculate two scenarios:
'   - What the CAO computed (agency's calculation)
'   - What QC determines is correct (QC's calculation)
'   
'   By having both columns, the examiner can compare and then select
'   which one represents the correct/final determination.
'
' DEPENDENCIES:
'   - Review_TANF_Utils.TANFfinalresults: Copies selected column to schedule
'   - TANF Computation sheet: Cell AL77 stores the selection
'
' CHANGE LOG:
'   2026-01-03  Created from UserForm1.frm
'               Added V2 comments, renamed to UF_TANF_ResultsColumn
' ============================================================================

Option Explicit

' ----------------------------------------------------------------------------
' Event: CommandButton1_Click (OK Button)
' ----------------------------------------------------------------------------
' PURPOSE:
'   Stores the column selection and calls the results transfer function.
' ----------------------------------------------------------------------------
Private Sub CommandButton1_Click()
    ' Validate that a column is selected
    If UF_TANF_ResultsColumn.ComboBox1.Value = "" Then
        MsgBox "Please pick a column to place your results.", vbExclamation
        UF_TANF_ResultsColumn.Hide
        Call redisplayform1
        Exit Sub
    End If
    
    ' Store selection in TANF Computation sheet
    Worksheets("TANF Computation").Range("AL77").Value = UF_TANF_ResultsColumn.ComboBox1.Value
    
    UF_TANF_ResultsColumn.Hide
    
    ' Transfer results from selected column to schedule
    ' V2: Call the renamed function in Review_TANF_Utils
    Call Review_TANF_Utils.TANFFinalResults
End Sub

' ----------------------------------------------------------------------------
' Event: CommandButton2_Click (Cancel Button)
' ----------------------------------------------------------------------------
Private Sub CommandButton2_Click()
    UF_TANF_ResultsColumn.Hide
    End
End Sub

' ----------------------------------------------------------------------------
' Helper: redisplayform1
' ----------------------------------------------------------------------------
Private Sub redisplayform1()
    UF_TANF_ResultsColumn.Show
End Sub


