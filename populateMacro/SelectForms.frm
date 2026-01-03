VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelectForms 
   Caption         =   "Select Form"
   ClientHeight    =   3192
   ClientLeft      =   40
   ClientTop       =   180
   ClientWidth     =   7520
   OleObjectBlob   =   "SelectForms.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SelectForms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




 'Userform code:
Private Sub CommandButton1_Click()
    Dim i As Integer, sht As String
    
Select Case Left(ActiveSheet.Name, 1)
    'SNAP Positive, SNAP Supplemental
    Case "5"
    For i = 0 To ListBox1.ListCount - 1
        If ListBox1.Selected(i) = True Then
            Unload SelectForms
            If i = 0 Then
               Call Finding_Memo_sub
            ElseIf i = 1 Then
                Call CAOAppt
            ElseIf i = 2 Then
                Call Rush
            ElseIf i = 3 Then
                Call TeleAppt
            ElseIf i = 4 Then
                Call QC14F
            ElseIf i = 5 Then
                Call QC14R
            ElseIf i = 6 Then
                Call QC14C
            ElseIf i = 7 Then
                Call timeliness_switch
            ElseIf i = 8 Then
                Call Info 'find code in CashMemos module
            ElseIf i = 9 Then
                Call Post
            ElseIf i = 10 Then
                Call SNAPPend
            ElseIf i = 11 Then
                Call SNAPDrop
            ElseIf i = 12 Then
                Call SNAPCAORequest
            ElseIf i = 13 Then
                Call Threshold
            ElseIf i = 14 Then
                Call LEP
            ElseIf i = 15 Then
                Call SpCAOAppt
            ElseIf i = 16 Then
                Call SpTeleAppt
            ElseIf i = 17 Then
                Call SNAPPA78
            ElseIf i = 18 Then
                Call SpPend
            End If
        End If
    Next i    'SNAP Negative
    Case "6"
    For i = 0 To ListBox1.ListCount - 1
        If ListBox1.Selected(i) = True Then
            Unload SelectForms
            If i = 0 Then
               Call Finding_Memo_sub
            ElseIf i = 1 Then
                Call Rush
            ElseIf i = 2 Then
                Call NewNeg ' Info Memo
            ElseIf i = 3 Then
                Call LEP
            End If
        End If
    Next i
    'MA Negative
    Case "8"
    For i = 0 To ListBox1.ListCount - 1
        If ListBox1.Selected(i) = True Then
            Unload SelectForms
            If i = 0 Then
               Call MA_Finding_Memo_sub("MA_Neg_Find")
            ElseIf i = 1 Then
               Call MA_Finding_Memo_sub("MA_Neg_Def")
            ElseIf i = 2 Then
                 Call PotentialErrorCall
            ElseIf i = 3 Then
                Call Info 'find code in CashMemos module
               'Call MA_Finding_Memo_sub("MA_Neg_Info")
            ElseIf i = 4 Then
                Call LEP
            End If
        End If
    Next i
    'MA Positive
    Case "2"
    
    If Left(ActiveSheet.Name, 2) = "24" Then 'PE Cases
      For i = 0 To ListBox1.ListCount - 1
        If ListBox1.Selected(i) = True Then
            Unload SelectForms
            If i = 0 Then
               Call MA_Finding_Memo_sub("MA_PE_Find")
            ElseIf i = 1 Then
                Call Info 'find code in CashMemos module
                'Call MA_Finding_Memo_sub("MA_PE_Info")
            'ElseIf i = 2 Then
             '   Call PE_Deficiency_Letter
            End If
        End If
    Next i
    Else 'MA Positive
        For i = 0 To ListBox1.ListCount - 1
        If ListBox1.Selected(i) = True Then
            Unload SelectForms
            If i = 0 Then
               Call MA_Finding_Memo_sub("MA_Pos_Find")
            ElseIf i = 1 Then
               Call MA_Finding_Memo_sub("MA_Pos_Def")
            ElseIf i = 2 Then
                Call MAAppt
            ElseIf i = 3 Then
                 Call ComSpouse
            ElseIf i = 4 Then
                 Call PotentialErrorCall
            ElseIf i = 5 Then
                Call Info 'find code in CashMemos module
                'Call MA_Finding_Memo_sub("MA_Pos_Info")
            ElseIf i = 6 Then
                Call QC14C
            ElseIf i = 7 Then
                Call QC14
            ElseIf i = 8 Then
                Call QC15
            ElseIf i = 9 Then
                Call Prelim_Info
            ElseIf i = 10 Then
                Call NHLLR
            ElseIf i = 11 Then
                Call NHBUS
            ElseIf i = 12 Then
                Call PendLet '472 Pending Letter
            ElseIf i = 13 Then
                Call SelfEmp ' PA472-Self Employment very basic
            ElseIf i = 14 Then
                Call SelfEmpDet 'more detailed self employment form
            ElseIf i = 15 Then
                Call PA76
            ElseIf i = 16 Then
                Call PA78
            ElseIf i = 17 Then
                Call PA83Z
            ElseIf i = 18 Then
                Call HouseholdComp
            ElseIf i = 19 Then
                Call MA_Finding_Memo_sub("MA_Pos_SAVE")
            ElseIf i = 20 Then
                Call MA_Supp
            ElseIf i = 21 Then
                Call MA_WAIVER
            ElseIf i = 22 Then
                Call Adoption
            ElseIf i = 23 Then
                Call MA_Zero
            ElseIf i = 24 Then
                Call MA_LTC_WAIVER
            ElseIf i = 25 Then
                Call Funeral
            ElseIf i = 26 Then
                Call LEP
            'End If
            ElseIf i = 27 Then
                Call MASpCAOAppt
            End If
        End If
    Next i
    End If
    'TANF
    Case "1"
    For i = 0 To ListBox1.ListCount - 1
        If ListBox1.Selected(i) = True Then
            Unload SelectForms
            If i = 0 Then
               Call Finding_Memo_sub
            ElseIf i = 1 Then
                Call Info 'find code in CashMemos module
            ElseIf i = 2 Then
                 Call PotentialErrorCall
            ElseIf i = 3 Then
                 Call TANF_Signature_Notification(1)
            'ElseIf i = 4 Then
            '     Call TANF_Signature_Notification(0)
            ElseIf i = 4 Then
                Call AMR
            ElseIf i = 5 Then
                Call Criminal
            ElseIf i = 6 Then
                Call TANFCAORequest
            ElseIf i = 7 Then
                Call TANF_SAVE
            ElseIf i = 8 Then
                Call TANFPA78
            ElseIf i = 9 Then
                Call LEP
            ElseIf i = 10 Then
                Call TANFPend
            ElseIf i = 11 Then
                Call TANFSchool
            End If
        End If
    Next i
    'GA
    Case "9"
    For i = 0 To ListBox1.ListCount - 1
        If ListBox1.Selected(i) = True Then
            Unload SelectForms
            If i = 0 Then
               Call Finding_Memo_sub
            ElseIf i = 1 Then
                Call AMR
            ElseIf i = 2 Then
                Call Criminal
            ElseIf i = 3 Then
                Call Info 'find code in CashMemos module
            ElseIf i = 4 Then
                 Call PotentialErrorCall
           ' ElseIf i = 5 Then
           '     Call TANFCAORequest
            ElseIf i = 5 Then
                Call TANF_SAVE
            ElseIf i = 6 Then
                Call TANFPA78
                
            End If
        End If
    Next i
End Select    'Find selected memo/form

    
End Sub
 
Private Sub CommandButton2_Click()
    Unload SelectForms
    End
End Sub
 
Private Sub ListBox1_Click()

End Sub

Private Sub UserForm_Initialize()
Select Case Left(ActiveSheet.Name, 1)
    'SNAP Positive, SNAP Supplemental
    Case "5"
            'List of memos/forms to select
            ListBox1.AddItem ("Findings Memo")
            ListBox1.AddItem ("CAO Appointment Letter")
            ListBox1.AddItem ("Case Summary Template")
            ListBox1.AddItem ("Telephone Appointment Letter")
            ListBox1.AddItem ("SNAP QC14F Memo")
            ListBox1.AddItem ("SNAP QC14R Memo")
            ListBox1.AddItem ("SNAP QC14C Memo")
            ListBox1.AddItem ("Timeliness Information/Finding Memo")
            ListBox1.AddItem ("Information Memo")
            ListBox1.AddItem ("Post Office Memo")
            ListBox1.AddItem ("Pending Letter")
            ListBox1.AddItem ("Drop Worksheet")
            ListBox1.AddItem ("CAO Forms Request")
            ListBox1.AddItem ("Error Under Threshold Memo")
            ListBox1.AddItem ("LEP")
            ListBox1.AddItem ("Spanish CAO Appointment Letter")
            ListBox1.AddItem ("Spanish Telephone Appointment Letter")
            ListBox1.AddItem ("PA78")
            ListBox1.AddItem ("Spanish Pending Letter")
            
    'SNAP Negative
    Case "6"
            ListBox1.AddItem ("Findings Memo")
            ListBox1.AddItem ("Case Summary Template")
            ListBox1.AddItem ("Negative Info Memo")
            ListBox1.AddItem ("LEP")
            
    'MA Negative
    Case "8"
            ListBox1.AddItem ("Findings Memo")
            ListBox1.AddItem ("Deficiency Memo")
            ListBox1.AddItem ("Potential Error Call Memo")
            ListBox1.AddItem ("Information Memo")
            ListBox1.AddItem ("LEP")
    'MA Positive
    Case "2"
        If Left(ActiveSheet.Name, 2) = "24" Then
            ListBox1.AddItem ("Findings Memo")
            ListBox1.AddItem ("Information Memo")
            'ListBox1.AddItem ("PE Deficiency Letter")
        Else
            ListBox1.AddItem ("Findings Memo")
            ListBox1.AddItem ("Deficiency Memo")
            ListBox1.AddItem ("MA Appointment Letter")
            ListBox1.AddItem ("Community Spouse Questionaire")
            ListBox1.AddItem ("Potential Error Call Memo")
            ListBox1.AddItem ("Information Memo")
            ListBox1.AddItem ("QC14 Coop Memo")
            ListBox1.AddItem ("QC14 CAO Request Memo")
            ListBox1.AddItem ("QC15 Memo")
            ListBox1.AddItem ("Preliminary Information Memo")
            ListBox1.AddItem ("NH LRR")
            ListBox1.AddItem ("NH Home Business")
            ListBox1.AddItem ("PA472-Pending Letter")
            ListBox1.AddItem ("PA472-Self Emp.")
            ListBox1.AddItem ("Self Emp.")
            ListBox1.AddItem ("PA76")
            ListBox1.AddItem ("PA78")
            ListBox1.AddItem ("PA83-Z")
            ListBox1.AddItem ("Household Composition")
            ListBox1.AddItem ("SAVE Deficiency Memo")
            ListBox1.AddItem ("MA Support Form")
            ListBox1.AddItem ("MA Waiver Memo")
            ListBox1.AddItem ("Adoption Foster Care")
            ListBox1.AddItem ("Zero Income Request")
            ListBox1.AddItem ("QC14 LTC Waiver Memo")
            ListBox1.AddItem ("Funeral Home Letter")
            ListBox1.AddItem ("LEP")
            ListBox1.AddItem ("Spanish MA App. Letter")
        End If
    'TANF
    Case "1"
            ListBox1.AddItem ("Findings Memo")
            ListBox1.AddItem ("Information Memo")
            ListBox1.AddItem ("Potential Error Call Memo")
            'ListBox1.AddItem ("Notification Signature Requirement Info Memo")
            ListBox1.AddItem ("Notification Requirement Info Memo")
            ListBox1.AddItem ("AMR")
            ListBox1.AddItem ("Criminal History")
            ListBox1.AddItem ("TANF CAO Request Form")
            ListBox1.AddItem ("SAVE")
            ListBox1.AddItem ("Employment/Earnings Report")
            ListBox1.AddItem ("LEP")
            ListBox1.AddItem ("TANF Pending")
            ListBox1.AddItem ("School Verification")
    'GA
    Case "9"
            ListBox1.AddItem ("Findings Memo")
            ListBox1.AddItem ("AMR Memo")
            ListBox1.AddItem ("Criminal History Memo")
            ListBox1.AddItem ("Information Memo")
            ListBox1.AddItem ("Potential Error Call Memo")
           ' ListBox1.AddItem ("GA CAO Request Form")
            ListBox1.AddItem ("SAVE")
            ListBox1.AddItem ("Employment/Earnings Report")
            
End Select
           
End Sub


