VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Where do Results Go?"
   ClientHeight    =   3210
   ClientLeft      =   80
   ClientTop       =   450
   ClientWidth     =   5880
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub CommandButton2_Click()
UserForm1.Hide
End


End Sub


Private Sub CommandButton1_Click()

If UserForm1.ComboBox1.Value = "" Then
    MsgBox "Please pick a column to place your results."
    UserForm1.Hide
    redisplayform1
End If

    Worksheets("TANF Computation").Range("AL77") = UserForm1.ComboBox1.Value
    UserForm1.Hide
    TANFfinalresults

End Sub

Private Sub UserForm_Click()

End Sub
