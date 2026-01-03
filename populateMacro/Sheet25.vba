VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet25"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Activate()

End Sub

Private Sub Worksheet_Change(ByVal Target As Range)



'David McRitchie, 2004-09-26, programming,  Case -- Entire Row
  '  http://www.mvps.org/dmcritchie/excel/event.htm#case
  
  Dim i As Long, j As Long
  Dim startrow As Long, endrow As Long, lastrow As Long
  'MsgBox Target.Row & " " & Target.Column
  
 Application.EnableEvents = True   'should be part of Change macro
If Target.Row = 6 And Target.Column = 66 Then

    If (Target.Value = "SAR Suspended Philly") Then
        ActiveSheet.Shapes("Text Box 18").TextFrame.Characters.Text = _
            "The action that was sampled was the (MM/DD/YY) suspension. " & Range("C8") & " was receiving SNAP benefits and was enrolled in Semi-Annual Reporting (SAR). The SAR from was mailed on (MM/DD/YY) with a SAR due date of (MM/DD/YY). The SAR form was not received by the SAR due date as required. As a result, a Late Incomplete Notice (LIN) was mailed on (MM/DD/YY) with a sanction override deadline of (MM/DD/YY).  This is a Philadelphia County case and since the completed SAR form or LIN was not received by the sanction override deadline the SNAP benefits were manually suspended on MM/DD/YY, effective MM/DD/YY. Quality Control (QC) determined that the MM/DD/YY manual suspension was valid for this Philadelphia County SNAP benefit.  According to Pennsylvania policy, Philadelphia County is required to manually suspend SNAP benefits that fail to comply with SAR reporting requirements timely in order to provide a safeguard against benefits suspending inappropriately."
    ElseIf (Target.Value = "SAR terminated NOT Philly") Then
        ActiveSheet.Shapes("Text Box 18").TextFrame.Characters.Text = "The action that was sampled was the (MM/DD/YY) termination.  " & Range("C8") & " was receiving SNAP benefits and was enrolled in Semi-Annual Reporting (SAR). The SAR from was mailed on (MM/DD/YY) with a SAR due date of (MM/DD/YY). The SAR form was not received by the SAR due date as required. As a result, a Late Incomplete Notice (LIN) was mailed on (MM/DD/YY) with a sanction override deadline of (MM/DD/YY). As the SAR form or LIN was not received by the sanction override deadline the SNAP benefits were automatically terminated on (MM/DD/YY), effective (MM/DD/YY). The SAR/LIN serves as advance notice of the termination action. Quality Control (QC) determined that the (MM/DD/YY) termination was valid."
    ElseIf (Target.Value = "Rejection 047 EX FS Issued") Then
         ActiveSheet.Shapes("Text Box 18").TextFrame.Characters.Text = "The action that was sampled was the (MM/DD/YY) rejection. A (Compass/walk-in/drop off) SNAP application was received on (MM/DD/YY) for " & Range("C8") & " applying for a (#) person SNAP household." _
         & vbNewLine & vbNewLine & "On MM/DD/YY the CAO authorized expedited SNAP benefits covering the time period MM/DD/YY through MM/DD/YY and they were available on MM/DD/YY (the ##th day pending from the application filing date).  IF NECESSARY USE: Quality Control (QC) determined that the expedited issuance was untimely as it was not available to the applicant within the allowable federal timeframe. As the expedited issuance was untimely the rejection action is considered invalid." _
         & vbNewLine & vbNewLine & "On (MM/DD/YY) the CAO sent an appointment letter and verification checklist to the applicant for their scheduled application interview on (MM/DD/YY) (at 00:00) or (between 00:00 and 00:00). The client failed to complete their scheduled application interview. As a result, the Notice of Missed Interview (NOMI) was mailed on (MM/DD/YY). On (MM/DD/YY), the 30th day pending, the CAO rejected the SNAP application to reason code 047 for failure to be interviewed as required. Quality Control (QC) determined that this rejection to reason code 047 was valid and proper notice was sent."
     ElseIf (Target.Value = "Rejection 047 EX FS Denied") Then
         ActiveSheet.Shapes("Text Box 18").TextFrame.Characters.Text = "The action that was sampled was the (MM/DD/YY) rejection. A (Compass/walk-in/drop off) SNAP application was received on (MM/DD/YY) for " & Range("C8") & " applying for a (#) person SNAP household." _
         & vbNewLine & vbNewLine & "The application listed shelter expenses of ($) monthly (rent/mortgage) as well as eligibility for the heating SUA ($536) as the applicant pays for their heating and/or cooling costs for total listed shelter/utility expenses of ($).  The application listed household income as follows: (list name, type and amount of income for each household member here) for total listed household income of ($) and resources of ($). On (MM/DD/YY) the CAO denied expedited SNAP benefits as listed household income and resources exceeded listed shelter/utility expenses.  QC determined that this expedited denial was correct." _
         & vbNewLine & vbNewLine & "On (MM/DD/YY) the CAO sent an appointment letter and verification checklist to the applicant for their scheduled application interview on (MM/DD/YY) (at 00:00) or (between 00:00 and 00:00). The client failed to complete their scheduled application interview. As a result, the Notice of Missed Interview (NOMI) was mailed on (MM/DD/YY). On (MM/DD/YY), the 30th day pending, the CAO rejected the SNAP application to reason code 047 for failure to be interviewed as required. Quality Control (QC) determined that this rejection to reason code 047 was valid and proper notice was sent."
    ElseIf (Target.Value = "Rejection 042 EX FS Issued") Then
        ActiveSheet.Shapes("Text Box 18").TextFrame.Characters.Text = "The action that was sampled was the (MM/DD/YY) rejection. A (Compass/walk-in/drop off) SNAP application was received on (MM/DD/YY) for " & Range("C8") & " applying for a (#) person SNAP household." _
        & vbNewLine & vbNewLine & "On MM/DD/YY the CAO authorized expedited SNAP benefits covering the time period MM/DD/YY through MM/DD/YY and they were available on MM/DD/YY (the ##th day pending from the application filing date).  IF NECESSARY USE: Quality Control (QC) determined that the expedited issuance was untimely as it was not available to the applicant within the allowable federal timeframe. As the expedited issuance was untimely the rejection action is considered invalid." _
        & vbNewLine & vbNewLine & "On (MM/DD/YY) the CAO sent an appointment letter and verification checklist to the applicant for their scheduled application interview on (MM/DD/YY) (at 00:00) or (between 00:00 and 00:00). The application interview was completed on (MM/DD/YY). On the verification checklist the CAO requested verification of: (list items requested). The applicant failed to provide the required verification as requested. As a result on (MM/DD/YY), the 30th day pending, the CAO rejected the SNAP application to reason code 042 for failure to provide: (list all items listed on the 042 rejection notice). Quality Control (QC) determined that this rejection to reason code 042 was valid and proper notice was sent."
    ElseIf (Target.Value = "Rejection 042 EX FS Denied") Then
         ActiveSheet.Shapes("Text Box 18").TextFrame.Characters.Text = "The action that was sampled was the (MM/DD/YY) rejection. A (Compass/walk-in/drop off) SNAP application was received on (MM/DD/YY) for " & Range("C8") & " applying for a (#) person SNAP household." _
        & vbNewLine & vbNewLine & "The application listed shelter expenses of ($) monthly (rent/mortgage) as well as eligibility for the heating SUA ($536) as the applicant pays for their heating and/or cooling costs for total listed shelter/utility expenses of ($). The application listed household income as follows: (list name, type and amount of income for each household member here) for total listed household income of ($) and resources of ($). On (MM/DD/YY) the CAO denied expedited SNAP benefits as listed household income and resources exceeded listed shelter/utility expenses." _
        & vbNewLine & vbNewLine & "On (MM/DD/YY) the CAO sent an appointment letter and verification checklist to the applicant for their scheduled application interview on (MM/DD/YY) (at 00:00) or (between 00:00 and 00:00). The application interview was completed on (MM/DD/YY). On the verification checklist the CAO requested verification of: (list items requested). The applicant failed to provide the required verification as requested. As a result on (MM/DD/YY), the 30th day pending, the CAO rejected the SNAP application to reason code 042 for failure to provide: (list all items listed on the 042 rejection notice). Quality Control (QC) determined that this rejection to reason code 042 was valid and proper notice was sent."
    ElseIf (Target.Value = "Valid SAR Termination of an Incomplete SAR form for All Counties Except Philadephia County") Then
        ActiveSheet.Shapes("Text Box 18").TextFrame.Characters.Text = _
            "The action that was sampled was the (MM/DD/YY) termination.  " & Range("C8") & " received SNAP benefits and was enrolled in Semi-Annual Reporting (SAR). The SAR form was mailed on (MM/DD/YY) with a SAR due date of (MM/DD/YY).  The SAR form was received on (MM/DD/YY) (prior to the SAR due date) but it was not complete as questions #(s) was/were not answered. The SAR form was tracked as incomplete.  As a result, on (MM/DD/YY) a Late/Incomplete Notice (LIN) was mailed indicating question(s) # was not answered with a sanction override deadline of (MM/DD/YY).  The customer failed to return a completed SAR form or LIN by the sanction override deadline therefore, the SNAP benefits were automatically terminated on (MM/DD/YY), effective (MM/DD/YY).  The LIN serves as advance notice of the suspension. Quality Control (QC) determined that the action to terminate the SNAP benefits was valid."
   
    End If
End If
'Clear schedule if case is a drop
If Target.Row = 29 And Target.Column = 6 Then
    If (Target.Value = 2 Or Target.Value = 3) Then
        Application.EnableEvents = False   'should be part of Change macro
  
        Range("m29") = "" '11
        Range("w29") = "" '12a
        Range("ag29") = "" '12b
        Range("al29") = "" '13
        Range("h33") = "" '14a
        Range("t33") = "" '14b
        Range("aa33") = "" '14c
        Range("al33") = "" '14d
        Range("j37") = "" '15a
        Range("r37") = "" '15b
        Range("aa37") = "" '15c
        Range("ai37") = "" '15d
        Range("h41") = "" '16a
        Range("s41") = "" '16b
        Range("aa41") = "" '16c
        Range("ak41") = "" '16d
        Range("e52") = "" '19
        Range("n52") = "" '20
        
    End If
 
'If Element code is entered, create comment and data validation for just this Element code
ElseIf Target.Row = 47 Then
  If (Target.Column = 5 Or Target.Column = 11 Or Target.Column = 17) And Target.Value <> "" Then
    Application.EnableEvents = False   'should be part of Change macro
    startrow = 0
    endrow = 0
    lastrow& = Range("BF" & Rows.Count).End(xlUp).Row
    'Find rows with data for this Element code
    For i = 2 To lastrow + 1
        If Target.Value = Range("BF" & i) And startrow = 0 Then
            startrow = i
        ElseIf Target.Value <> Range("BF" & i) And startrow <> 0 Then
            endrow = i - 1
            Exit For
        End If
    Next i
    'MsgBox startrow & " " & endrow
    
    'Create comment
    TempStr = ""
    For j = startrow To endrow
        TempStr = TempStr & Range("BH" & j) & vbCrLf
    Next j
    Cells(47, Target.Column + 18).Comment.Text TempStr
    
    'Create data validation
    TempStr = "=$BG$" & startrow & ":$BG$" & endrow
    Cells(47, Target.Column + 18).Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=TempStr
        .IgnoreBlank = True
        .InCellDropdown = False
        .InputTitle = ""
        .ErrorTitle = "Nature"
        .InputMessage = ""
        .ErrorMessage = _
        "Please enter a valid Nature code.  Refer to comment for description of codes."
        .ShowInput = True
        .ShowError = True
    End With

  End If

End If

    Application.EnableEvents = True   'should be part of Change macro

End Sub

'Sub SNAPNarrative()
'If Range("BN6") = "SAR Suspended Philly" Then

'    ActiveSheet.Shapes("Text Box 17").TextFrame.Characters.Text = "This is a Philadelphia County case and since the completed SAR form or LIN was not received by the sanction override deadline the SNAP benefits were manually suspended on MM/DD/YY, effective MM/DD/YY. Quality Control (QC) determined that the MM/DD/YY manual suspension was valid for this Philadelphia County SNAP benefit.  According to Pennsylvania policy, Philadelphia County is required to manually suspend SNAP benefits that fail to comply with SAR reporting requirements timely in order to provide a safeguard against benefits suspending inappropriately."
'ElseIf Range("BN6") = "SAR not terminated Philly" Then
'    ActiveSheet.Shapes("Text Box 17").TextFrame.Characters.Text = " This is a Philadelphia County case and since the completed SAR form or LIN was not received by the sanction override deadline the SNAP benefits were manually suspended on MM/DD/YY, effective MM/DD/YY. Quality Control (QC) determined that the MM/DD/YY manual suspension was valid for this Philadelphia County SNAP benefit.  According to Pennsylvania policy, Philadelphia County is required to manually suspend SNAP benefits that fail to comply with SAR reporting requirements timely in order to provide a safeguard against benefits suspending inappropriately."
'ElseIf Range("BN6") = "Rejection 047" Then
'    ActiveSheet.Shapes("Text Box 17").TextFrame.Characters.Text = "On (MM/DD/YY) the CAO sent an appointment letter and verification checklist to the applicant for their scheduled application interview on (MM/DD/YY) (at 00:00) or (between 00:00 and 00:00). The client failed to complete their scheduled application interview. As a result, the Notice of Missed Interview (NOMI) was mailed on (MM/DD/YY). On (MM/DD/YY), the 30th day pending, the CAO rejected the SNAP application to reason code 047 for failure to be interviewed as required. Quality Control (QC) determined that this rejection to reason code 047 was valid and proper notice was sent. "
'ElseIf Range("BN6") = "Rejection 042" Then
'    ActiveSheet.Shapes("Text Box 17").TextFrame.Characters.Text = "On (MM/DD/YY) the CAO sent an appointment letter and verification checklist to the applicant for their scheduled application interview on (MM/DD/YY) (at 00:00) or (between 00:00 and 00:00). The application interview was completed on (MM/DD/YY). On the verification checklist the CAO requested verification of: (list items requested). The applicant failed to provide the required verification as requested. As a result on (MM/DD/YY), the 30th day pending, the CAO rejected the SNAP application to reason code 042 for failure to provide: (list all items listed on the 042 rejection notice). Quality Control (QC) determined that this rejection to reason code 042 was valid and proper notice was sent."

'End If
'Select Case Narrative
'    Case "SAR Suspended Philly"
'        ActiveSheet.Shapes("Text Box 17").TextFrame.Characters.Text = "This is a Philadelphia County case and since the completed SAR form or LIN was not received by the sanction override deadline the SNAP benefits were manually suspended on MM/DD/YY, effective MM/DD/YY. Quality Control (QC) determined that the MM/DD/YY manual suspension was valid for this Philadelphia County SNAP benefit.  According to Pennsylvania policy, Philadelphia County is required to manually suspend SNAP benefits that fail to comply with SAR reporting requirements timely in order to provide a safeguard against benefits suspending inappropriately."
'    Case "SAR not terminated Philly"
'        ActiveSheet.Shapes("Text Box 17").TextFrame.Characters.Text = " This is a Philadelphia County case and since the completed SAR form or LIN was not received by the sanction override deadline the SNAP benefits were manually suspended on MM/DD/YY, effective MM/DD/YY. Quality Control (QC) determined that the MM/DD/YY manual suspension was valid for this Philadelphia County SNAP benefit.  According to Pennsylvania policy, Philadelphia County is required to manually suspend SNAP benefits that fail to comply with SAR reporting requirements timely in order to provide a safeguard against benefits suspending inappropriately."
'    Case "Rejection 047"
'        ActiveSheet.Shapes("Text Box 17").TextFrame.Characters.Text = "On (MM/DD/YY) the CAO sent an appointment letter and verification checklist to the applicant for their scheduled application interview on (MM/DD/YY) (at 00:00) or (between 00:00 and 00:00). The client failed to complete their scheduled application interview. As a result, the Notice of Missed Interview (NOMI) was mailed on (MM/DD/YY). On (MM/DD/YY), the 30th day pending, the CAO rejected the SNAP application to reason code 047 for failure to be interviewed as required. Quality Control (QC) determined that this rejection to reason code 047 was valid and proper notice was sent. "
'    Case "Rejection 042"
'        ActiveSheet.Shapes("Text Box 17").TextFrame.Characters.Text = "On (MM/DD/YY) the CAO sent an appointment letter and verification checklist to the applicant for their scheduled application interview on (MM/DD/YY) (at 00:00) or (between 00:00 and 00:00). The application interview was completed on (MM/DD/YY). On the verification checklist the CAO requested verification of: (list items requested). The applicant failed to provide the required verification as requested. As a result on (MM/DD/YY), the 30th day pending, the CAO rejected the SNAP application to reason code 042 for failure to provide: (list all items listed on the 042 rejection notice). Quality Control (QC) determined that this rejection to reason code 042 was valid and proper notice was sent."
'    End Select

'End Sub


