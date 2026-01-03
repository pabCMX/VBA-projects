VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_TANF_Helper 
   Caption         =   "TANF Helper"
   ClientHeight    =   3500
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4500
   OleObjectBlob   =   "UF_TANF_Helper.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_TANF_Helper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ============================================================================
' UF_TANF_Helper - TANF Computation Helper Form
' ============================================================================
' WHAT THIS FORM DOES:
'   Provides additional helper functions for TANF computations. Assists
'   with data entry and validation on the computation sheet.
'
' CHANGE LOG:
'   2026-01-02  Renamed from UserForm2 - added V2 comments
' ============================================================================

Option Explicit

Private Sub UserForm_Initialize()
    ' Initialize form
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub


