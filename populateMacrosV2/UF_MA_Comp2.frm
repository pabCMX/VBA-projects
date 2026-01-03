VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_MA_Comp2 
   Caption         =   "Final SACQ"
   ClientHeight    =   3225
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   4710
   OleObjectBlob   =   "UF_MA_Comp2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_MA_Comp2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ============================================================================
' UF_MA_Comp2 - MA Computation Helper Form 2 (SACQ Column Selector)
' ============================================================================
' WHAT THIS FORM DOES:
'   Second computation helper for MA Positive reviews. Allows the examiner
'   to select which column from the MA computation sheet contains the final
'   SACQ (Spenddown Amount Counted in QC) values to be transferred.
'
' HOW IT WORKS:
'   1. User selects a column (D through M) from ComboBox1
'   2. The column letter is extracted and stored in module variable columnletter
'   3. Calls MA_Comp_finalresults to copy values from selected column
'
' FORM CONTROLS:
'   - ComboBox1: Column selector (D, E, F, H, I, J, K, L, M)
'   - CommandButton1: OK - Accept the selection
'   - CommandButton2: Cancel - Close without action
'
' OUTPUT:
'   - Sets module-level variable columnletter (e.g., "D", "E", etc.)
'   - Calls MA_Comp_finalresults to process the transfer
'
' WHY MULTIPLE COLUMNS:
'   The MA Computation sheet allows for multiple scenarios to be calculated
'   side-by-side. The examiner selects which scenario is correct.
'
' CHANGE LOG:
'   2026-01-03  Renamed from UserFormMAC2 to UF_MA_Comp2
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
    If UF_MA_Comp2.ComboBox1.Value = "" Then
        MsgBox "Please pick a column to place your results."
        UF_MA_Comp2.Hide
        Call redisplayformMAC2
        Exit Sub
    End If
    
    ' Extract the column letter from the selection (e.g., "Column D" -> "D")
    ' V2: Use the module-level variable in Review_MA_Comp
    Review_MA_Comp.MACompColumnLetter = Right(UF_MA_Comp2.ComboBox1.Value, 1)
    
    ' Clear the selection for next use
    UF_MA_Comp2.ComboBox1.Value = ""
    
    UF_MA_Comp2.Hide
    
    ' Call the results transfer function
    ' V2: Call the renamed function in Review_MA_Comp
    Call Review_MA_Comp.MACompFinalResults
End Sub

' ----------------------------------------------------------------------------
' Event: CommandButton2_Click (Cancel Button)
' ----------------------------------------------------------------------------
' PURPOSE:
'   Closes the form without processing.
' ----------------------------------------------------------------------------
Private Sub CommandButton2_Click()
    UF_MA_Comp2.Hide
    End
End Sub

' ----------------------------------------------------------------------------
' Event: UserForm_Initialize
' ----------------------------------------------------------------------------
' PURPOSE:
'   Populates the ComboBox with available column choices.
' ----------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    ' Initialize columns D, E, F, H-M (note: G is typically skipped)
    ComboBox1.AddItem ("Column D")
    ComboBox1.AddItem ("Column E")
    ComboBox1.AddItem ("Column F")
    ComboBox1.AddItem ("Column H")
    ComboBox1.AddItem ("Column I")
    ComboBox1.AddItem ("Column J")
    ComboBox1.AddItem ("Column K")
    ComboBox1.AddItem ("Column L")
    ComboBox1.AddItem ("Column M")
    ComboBox1.Value = ""
End Sub

