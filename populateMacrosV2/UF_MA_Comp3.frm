VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_MA_Comp3 
   Caption         =   "Final SACQ"
   ClientHeight    =   3225
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   4710
   OleObjectBlob   =   "UF_MA_Comp3.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_MA_Comp3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ============================================================================
' UF_MA_Comp3 - MA Computation Helper Form 3 (SACQ Column Selector)
' ============================================================================
' WHAT THIS FORM DOES:
'   Third computation helper for MA Positive reviews. Similar to UF_MA_Comp2
'   but calls a different results transfer function (MA_Comp_finalresults3).
'
' HOW IT WORKS:
'   1. User selects a column (D through M) from ComboBox1
'   2. The column letter is extracted and stored in module variable columnletter
'   3. Calls MA_Comp_finalresults3 to copy values from selected column
'
' FORM CONTROLS:
'   - ComboBox1: Column selector (D, E, F, H, I, J, K, L, M)
'   - CommandButton1: OK - Accept the selection
'   - CommandButton2: Cancel - Close without action
'
' OUTPUT:
'   - Sets module-level variable columnletter (e.g., "D", "E", etc.)
'   - Calls MA_Comp_finalresults3 to process the transfer
'
' DIFFERENCE FROM UF_MA_Comp2:
'   This form calls MA_Comp_finalresults3 instead of MA_Comp_finalresults,
'   which may transfer different fields or to different locations.
'
' CHANGE LOG:
'   2026-01-03  Renamed from UserFormMAC3 to UF_MA_Comp3
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
    If UF_MA_Comp3.ComboBox1.Value = "" Then
        MsgBox "Please pick a column to place your results."
        UF_MA_Comp3.Hide
        Call redisplayformMAC3
        Exit Sub
    End If
    
    ' Extract the column letter from the selection (e.g., "Column D" -> "D")
    ' V2: Use the module-level variable in Review_MA_Comp
    Review_MA_Comp.MACompColumnLetter = Right(UF_MA_Comp3.ComboBox1.Value, 1)
    
    ' Clear the selection for next use
    UF_MA_Comp3.ComboBox1.Value = ""
    
    UF_MA_Comp3.Hide
    
    ' Call the results transfer function (variant 3)
    ' V2: Call the renamed function in Review_MA_Comp
    Call Review_MA_Comp.MACompFinalResults3
End Sub

' ----------------------------------------------------------------------------
' Event: CommandButton2_Click (Cancel Button)
' ----------------------------------------------------------------------------
' PURPOSE:
'   Closes the form without processing.
' ----------------------------------------------------------------------------
Private Sub CommandButton2_Click()
    UF_MA_Comp3.Hide
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

