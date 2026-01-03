VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "Pick Drive"
   ClientHeight    =   4020
   ClientLeft      =   80
   ClientTop       =   440
   ClientWidth     =   5880
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton2_Click()
UserForm3.Hide
End
End Sub


Private Sub CommandButton1_Click()

If UserForm3.ComboBox1.Value = "" Then
    MsgBox "Please pick a drive where your DQC folder is."
    UserForm3.Hide
    redisplayform3
End If

    Worksheets("Populate").Range("S10") = UserForm3.ComboBox1.Value
    UserForm3.Hide
    
    UserForm50.Show

End Sub

Private Sub UserForm_Click()

End Sub
