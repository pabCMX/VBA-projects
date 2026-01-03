VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormMAC2 
   Caption         =   "Final SACQ"
   ClientHeight    =   3225
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   4710
   OleObjectBlob   =   "UserFormMAC2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormMAC2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CommandButton2_Click()
UserFormMAC2.Hide
End
End Sub

Private Sub CommandButton1_Click()

If UserFormMAC2.ComboBox1.Value = "" Then
    MsgBox "Please pick a column to place your results."
    UserFormMAC2.Hide
    redisplayformMAC2
End If

    columnletter = Right(UserFormMAC2.ComboBox1.Value, 1)
    UserFormMAC2.ComboBox1.Value = ""
    UserFormMAC2.Hide
    MA_Comp_finalresults

End Sub


Private Sub UserForm_Initialize()

    'Initialize columns
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
