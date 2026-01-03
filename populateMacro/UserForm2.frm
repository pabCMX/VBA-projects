VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "Final Determination"
   ClientHeight    =   3225
   ClientLeft      =   80
   ClientTop       =   440
   ClientWidth     =   5880
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CommandButton2_Click()
UserForm2.Hide
End
End Sub


Private Sub CommandButton1_Click()

If UserForm2.ComboBox1.Value = "" Then
    MsgBox "Please pick a column to place your results."
    UserForm2.Hide
    redisplayform2
End If

    Worksheets("TANF Computation").Range("AL78") = UserForm2.ComboBox1.Value
    UserForm2.Hide
    finaldetermination

End Sub


Private Sub UserForm_Click()

End Sub
