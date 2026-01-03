VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm50 
   Caption         =   "Populate Review Schedule"
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   4680
   OleObjectBlob   =   "UserForm50.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm50"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

If UserForm50.ComboBox1.Value = "" Or UserForm50.ComboBox2.Value = "" Then
    MsgBox "Please fill in both Program and Month fields"
    UserForm50.Hide
    redisplayform50
End If

If OptionButton1.Value = False And OptionButton2.Value = False Then
    MsgBox "Please select either Schedule or Transmittal"
    UserForm50.Hide
    redisplayform50
End If

Worksheets("Populate").Range("w7") = UserForm50.ComboBox1.Value
    Worksheets("Populate").Range("x7") = UserForm50.ComboBox2.Value
    UserForm50.Hide
    
    If OptionButton1.Value = True Then
       Review_Schedule
    ElseIf OptionButton2.Value = True Then
       Transmittal
    End If


End Sub

Private Sub CommandButton2_Click()
UserForm50.Hide
End

End Sub

Private Sub OptionButton2_Click()

End Sub

Private Sub UserForm_Click()

End Sub
