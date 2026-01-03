VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_TANF_FinalDetermination 
   Caption         =   "Final Determination"
   ClientHeight    =   3225
   ClientLeft      =   80
   ClientTop       =   440
   ClientWidth     =   5880
   OleObjectBlob   =   "UF_TANF_FinalDetermination.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_TANF_FinalDetermination"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ============================================================================
' UF_TANF_FinalDetermination - TANF Final Determination Column Selector
' ============================================================================
' WHAT THIS FORM DOES:
'   Similar to UF_TANF_ResultsColumn, but specifically for the final
'   determination step. After the examiner has completed their review
'   and comparison, this form is used to select which column contains
'   the final determination values.
'
' FORM CONTROLS:
'   - ComboBox1: Column selection (D or E)
'   - CommandButton1: OK - proceeds with final determination
'   - CommandButton2: Cancel - aborts
'
' DIFFERENCE FROM UF_TANF_ResultsColumn:
'   - ResultsColumn: Used during initial calculation phase
'   - FinalDetermination: Used at the end to set the official finding
'
' DEPENDENCIES:
'   - Review_TANF_Utils.finaldetermination: Processes the final determination
'   - TANF Computation sheet: Cell AL78 stores the selection
'
' CHANGE LOG:
'   2026-01-03  Created from UserForm2.frm
'               Added V2 comments, renamed to UF_TANF_FinalDetermination
' ============================================================================

Option Explicit

' ----------------------------------------------------------------------------
' Event: CommandButton1_Click (OK Button)
' ----------------------------------------------------------------------------
Private Sub CommandButton1_Click()
    ' Validate that a column is selected
    If UF_TANF_FinalDetermination.ComboBox1.Value = "" Then
        MsgBox "Please pick a column to place your results.", vbExclamation
        UF_TANF_FinalDetermination.Hide
        Call redisplayform2
        Exit Sub
    End If
    
    ' Store selection in TANF Computation sheet
    Worksheets("TANF Computation").Range("AL78").Value = UF_TANF_FinalDetermination.ComboBox1.Value
    
    UF_TANF_FinalDetermination.Hide
    
    ' Process final determination
    ' V2: Call the renamed function in Review_TANF_Utils
    Call Review_TANF_Utils.TANFFinalDetermination
End Sub

' ----------------------------------------------------------------------------
' Event: CommandButton2_Click (Cancel Button)
' ----------------------------------------------------------------------------
Private Sub CommandButton2_Click()
    UF_TANF_FinalDetermination.Hide
    End
End Sub

' ----------------------------------------------------------------------------
' Helper: redisplayform2
' ----------------------------------------------------------------------------
Private Sub redisplayform2()
    UF_TANF_FinalDetermination.Show
End Sub


