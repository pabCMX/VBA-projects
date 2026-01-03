VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_TANF_Results 
   Caption         =   "TANF Results"
   ClientHeight    =   4000
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5000
   OleObjectBlob   =   "UF_TANF_Results.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_TANF_Results"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ============================================================================
' UF_TANF_Results - TANF Final Results Helper Form
' ============================================================================
' WHAT THIS FORM DOES:
'   Helps examiners transfer final computation results to the schedule.
'   Displays a summary of the computation and allows selection of which
'   column's results to use.
'
' WHEN TO USE:
'   After completing computations on the TANF Computation sheet, use this
'   form to transfer the final results back to the main schedule.
'
' CONTROLS:
'   - optColD      : Option to use Column D results
'   - optColE      : Option to use Column E results
'   - lblSummary   : Shows computation summary
'   - cmdTransfer  : Transfer results to schedule
'   - cmdCancel    : Close without action
'
' CHANGE LOG:
'   2026-01-02  Renamed from UserForm1 - added V2 comments
' ============================================================================

Option Explicit

' ----------------------------------------------------------------------------
' Event: UserForm_Initialize
' ----------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    ' Load summary data from computation sheet
    On Error Resume Next
    
    Dim compSheet As Worksheet
    Set compSheet = ActiveWorkbook.Worksheets("TANF Computation")
    
    If Not compSheet Is Nothing Then
        ' Display summary info
        ' lblSummary.Caption = "Column D Amount: " & compSheet.Range("D62").Value & vbCrLf & _
        '                      "Column E Amount: " & compSheet.Range("E62").Value
    End If
    
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' Event: cmdTransfer_Click
' ----------------------------------------------------------------------------
Private Sub cmdTransfer_Click()
    ' Transfer results based on selection
    On Error Resume Next
    
    ' Call the transfer function
    Call finalresults
    
    Unload Me
    
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' Event: cmdCancel_Click
' ----------------------------------------------------------------------------
Private Sub cmdCancel_Click()
    Unload Me
End Sub


