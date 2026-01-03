VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelectDate 
   Caption         =   "Enter date for CAO Appointment"
   ClientHeight    =   3225
   ClientLeft      =   80
   ClientTop       =   440
   ClientWidth     =   5880
   OleObjectBlob   =   "SelectDate.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SelectDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
    date_array = Split(TextBox1.Value, "/")
    date_entered = DateSerial(date_array(2), date_array(0), date_array(1))
    If date_entered < Date Then
        MsgBox "Appointment date is in the past. Please choose another date."
    ElseIf Weekday(date_entered) = 1 Or Weekday(date_entered) = 7 Then
        MsgBox "Appointment date is on a weekend. Please choose another date."
    Else
        ApptDate = Format(date_entered, "dddd, mmmm dd, yyyy")
        Unload Me
    End If
End Sub


Private Sub CommandButton2_Click()
    Unload Me
End Sub


Private Sub UserForm_Activate()
    TextBox1.Value = Date
End Sub

