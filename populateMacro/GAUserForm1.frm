VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GAUserForm1 
   Caption         =   "Where do Results Go?"
   ClientHeight    =   3210
   ClientLeft      =   50
   ClientTop       =   350
   ClientWidth     =   4710
   OleObjectBlob   =   "GAUserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GAUserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton2_Click()
    GAUserForm1.Hide
    End

End Sub


Private Sub CommandButton1_Click()

  If GAUserForm1.ComboBox1.Value = "" Then
    MsgBox "Please pick a column to place your results."
    GAUserForm1.Hide
    redisplayGAform1
  End If

  Worksheets("GA Computation").Range("AL77") = GAUserForm1.ComboBox1.Value
  GAUserForm1.Hide
  GAfinalresults

End Sub


Private Sub UserForm_Click()

End Sub
