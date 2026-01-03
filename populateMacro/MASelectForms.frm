VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MASelectForms 
   Caption         =   "Select Form"
   ClientHeight    =   3225
   ClientLeft      =   80
   ClientTop       =   440
   ClientWidth     =   7210
   OleObjectBlob   =   "MASelectForms.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MASelectForms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ListBox1_Click()

End Sub

Private Sub UserForm_Click()

 'Userform code:
Private Sub CommandButton1_Click()
    Dim i As Integer, sht As String
    
    'Find selected memo/form
    For i = 0 To ListBox1.ListCount - 1
        If ListBox1.Selected(i) = True Then
            Unload MASelectForms
            If i = 0 Then
               Call MailMergeandSave
            ElseIf i = 1 Then
                Call Taxonomy
            'ElseIf i = 2 Then
            '    Call MAAppt
             ElseIf i = 2 Then
                 Call ComSpouse
            ElseIf i = 3 Then
                 Call QC14C
            ElseIf i = 4 Then
                 Call QC15
                
            End If
        End If
    Next i
    
End Sub
 
Private Sub CommandButton2_Click()
    Unload SelectForms
    End
End Sub
 
Private Sub UserForm_Initialize()
Select Case Left(ActiveSheet.Name, 1)
   
    'MA Negative
    Case "8"
            ListBox1.AddItem ("Findings Memo")
            ListBox1.AddItem ("Taxonomy Information Memo")
    'MA Positive
    Case "2"
            ListBox1.AddItem ("Findings Memo")
            ListBox1.AddItem ("Taxonomy Information Memo")
            'ListBox1.AddItem ("MA Appointment Letter")
            ListBox1.AddItem ("Community Spouse")
            ListBox1.AddItem ("QC 14")
            ListBox1.AddItem ("QC 15")
    
End Select

End Sub
