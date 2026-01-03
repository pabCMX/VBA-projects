VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelectTime 
   Caption         =   "UserForm3"
   ClientHeight    =   3225
   ClientLeft      =   80
   ClientTop       =   440
   ClientWidth     =   5880
   OleObjectBlob   =   "SelectTime.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SelectTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 'Userform code:
Private Sub CommandButton1_Click()
    Dim i As Integer, sht As String
    
    'Find selected time
    For i = 0 To ListBox1.ListCount - 1
        If ListBox1.Selected(i) = True Then
            ApptTime = ListBox1.Value
            Unload SelectTime
        End If
    Next i
    
End Sub
 
Private Sub CommandButton2_Click()
    Unload SelectTime
    End
End Sub
 
Private Sub UserForm_Initialize()

            'Initialize times
            ListBox1.AddItem ("8:00 AM")
            ListBox1.AddItem ("8:15 AM")
            ListBox1.AddItem ("8:30 AM")
            ListBox1.AddItem ("8:45 AM")
            ListBox1.AddItem ("9:00 AM")
            ListBox1.AddItem ("9:15 AM")
            ListBox1.AddItem ("9:30 AM")
            ListBox1.AddItem ("9:45 AM")
            ListBox1.AddItem ("10:00 AM")
            ListBox1.AddItem ("10:15 AM")
            ListBox1.AddItem ("10:30 AM")
            ListBox1.AddItem ("10:45 AM")
            ListBox1.AddItem ("11:00 AM")
            ListBox1.AddItem ("11:15 AM")
            ListBox1.AddItem ("11:30 AM")
            ListBox1.AddItem ("11:45 AM")
            ListBox1.AddItem ("12:00 PM")
            ListBox1.AddItem ("12:15 PM")
            ListBox1.AddItem ("12:30 PM")
            ListBox1.AddItem ("12:45 PM")
            ListBox1.AddItem ("1:00 PM")
            ListBox1.AddItem ("1:15 pM")
            ListBox1.AddItem ("1:30 PM")
            ListBox1.AddItem ("1:45 PM")
            ListBox1.AddItem ("2:00 PM")
            ListBox1.AddItem ("2:15 pM")
            ListBox1.AddItem ("2:30 PM")
            ListBox1.AddItem ("2:45 PM")
            ListBox1.AddItem ("3:00 PM")
            ListBox1.AddItem ("3:15 pM")
            ListBox1.AddItem ("3:30 PM")
            ListBox1.AddItem ("3:45 PM")
            ListBox1.AddItem ("4:00 PM")
           
End Sub
 


