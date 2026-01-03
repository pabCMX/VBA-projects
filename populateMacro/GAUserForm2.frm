VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GAUserForm2 
   Caption         =   "Final Determination"
   ClientHeight    =   3225
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   4710
   OleObjectBlob   =   "GAUserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GAUserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton2_Click()
    GAUserForm2.Hide
    End
End Sub


Private Sub CommandButton1_Click()

  If GAUserForm2.ComboBox1.Value = "" Then
    MsgBox "Please pick a column to place your results."
    GAUserForm2.Hide
    redisplayGAform2
  End If

  Worksheets("GA Computation").Range("AL78") = GAUserForm2.ComboBox1.Value
  GAUserForm2.Hide
  GAfinaldetermination

End Sub

Private Sub UserForm_Click()

End Sub
