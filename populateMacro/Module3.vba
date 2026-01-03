Attribute VB_Name = "Module3"
Sub snap_edit_check_pos()
'this routine looks for edits in the SNAP Pos Schedule
    Dim thisws As Worksheet
    Dim datasourcewb As Workbook, datasourcews As Worksheet
    Dim thiswb As Workbook
    
    Set thiswb = ActiveWorkbook
    Set thisws = ActiveSheet
    
    'check ABAWD status in Section 4
    If thisws.Range("C22") = 1 Then
    For irow = 89 To 122 Step 3
            ' If person number is blank then end of data
            If thisws.Range("B" & irow) = "" Then
                Exit For
            End If
            'Check if SNAP Participation is 1 and Age is not between 18 and 49, AWAWD must be 1
            If Val(thisws.Range("E" & irow)) = 1 And Val(thisws.Range("AH" & irow)) <> 9 _
                And Not (Val(thisws.Range("J" & irow)) > 17 And Val(thisws.Range("J" & irow)) < 50) _
                 Then
               MsgBox "In Section 4, Person Number " & thisws.Range("B" & irow) & " is participating in SNAP (Item 47 = 1) " & _
               "and is not between 18 and 49 years old (Item 49), but their ABAWD Status (Item 57) is not equal to 9."
               thisws.Range("AH" & irow).Select
               End
            End If
    Next irow
    End If
    
    'Check if there is an Element entered if this is an Error Case
    If thisws.Range("K22") > "1" And thisws.Range("B29") = "" Then
        MsgBox "This is an Error case (Item 8 is not equal to 1), but there is not an Element Code entered in Item 12."
        thisws.Range("B29").Select
        End
    End If
    
    'Check if there is a 1 in Item 23 and $ amt greater then 0 in Item 24
    If thisws.Range("AB50") = 1 And Not (thisws.Range("AI50") = "" Or thisws.Range("AI50") = "-") Then
        MsgBox "There is an no adjustment for allotment (Item 23), but there is a dollar amount greater then $0 in Item 24 ."
        thisws.Range("AB50").Select
        End
    End If
    
    'Check if there is a 1 in Item 44a. If there is, then 44b should be blank or dash.
    If thisws.Range("W82") = 1 And Not (thisws.Range("AA82") = "" Or thisws.Range("AA82") = "-") Then
        MsgBox "Item 44a is 1, but Item 44b is not blank or a dash."
        thisws.Range("AA82").Select
        End
    End If
    
    'check if there is >1 in item 44a. If there is, then 44b should have 1 or 2
    If thisws.Range("W82") > 1 And thisws.Range("AA82") = "" Then
        MsgBox "If item 44a is greater than 1 then item 44b needs to be a 1 or 2."
        End
    End If
    
    'LEP is blank in Section 7, Line 4, Box 2
    If thisws.Range("C159") = "" Then
        MsgBox "LEP in Section 7, Line 4, Box 2 is Blank.  Please enter a value"
        End
    End If
    
    'check Earned Incomes in Section 5 and compare to Employment Status in Section 4
    'Look in Section 5 for Income Types 11, 12, 14 and 16
    For irow = 131 To 143 Step 3
        'Check for end of data
        If thisws.Range("B" & irow) = "" Then
            Exit For
        End If
        'Check columns with income types
        For icol = 5 To 32 Step 9
            'Check for Income Types 11, 12, 14 and 16
            If Val(thisws.Cells(irow, icol)) = 11 Or Val(thisws.Cells(irow, icol)) = 12 Or _
                Val(thisws.Cells(irow, icol)) = 14 Or Val(thisws.Cells(irow, icol)) = 16 Then
                'set flag to see if person number in Section 5 was found in Section 4
                found_flag = 0
                'Check in Section 4 that same person Employment Status is not 1 or 2
                For jrow = 89 To 122 Step 3
                    'Look for person number in Section 4
                    If Val(thisws.Range("B" & jrow)) = Val(thisws.Range("B" & irow)) Then
                        found_flag = 1
                        If Val(thisws.Range("Y" & jrow)) = 1 Or Val(thisws.Range("Y" & jrow)) = 2 Then
                            MsgBox "Person number " & thisws.Range("B" & jrow) & " had earned income in " & _
                            "Section 5, but is indicated as not in labor force or unemployed in Section 4."
                            thisws.Range("Y" & jrow).Select
                            End
                        End If
                    End If
                Next jrow
                'If no matching person number was found in Section 4, send message
                If found_flag = 0 Then
                    MsgBox "Person number " & thisws.Range("B" & irow) & "  in " & _
                    "Section 5 could not be found in Section 4."
                    thisws.Range("B" & jrow).Select
                    End
                End If
                    
            End If
        Next icol
    Next irow
                
    'Message that all edits passed OK
    MsgBox "All edits passed. An email will be sent to your supervisor that this review is completed."
    'send email to supervisor
    Call send_email("SNAP Positive", thisws)
    'Put cwopa and date on schedule
    Range("AP13") = Environ("USERNAME") & " " & Date
    thiswb.Worksheets("FS Workbook").Range("G40") = Date
    
End Sub
Sub snap_return()
'this routine looks for edits in the SNAP Pos Schedule
    Dim thisws As Worksheet
    Dim datasourcewb As Workbook, datasourcews As Worksheet
    Dim thiswb As Workbook
    
    Set thiswb = ActiveWorkbook
    Set thisws = ActiveSheet
    'send email to supervisor
    Call send_email_edit("SNAP Positive", thisws)
    'Put cwopa and date on schedule
    Range("CB7") = Environ("USERNAME") & " " & Date
     
End Sub
Sub snap_edit_check_neg()
'this routine looks for edits in the SNAP Negative Schedule
    Dim thisws As Worksheet
    Dim datasourcewb As Workbook, datasourcews As Worksheet

    Set thisws = ActiveSheet
    
    'Check if case is valid and notice was sent then required language is 1
'    If thisws.Range("M29") = 1 And (thisws.Range("AL29") = 1 Or thisws.Range("AL29") = 3) And Not thisws.Range("H33") = 1 Then
'        MsgBox "Case is valid (Item 11) and notice was sent (Item 13), but the required language isn't 1 (Item 14a)"
'        thisws.Range("H33").Select
'        End
'    End If
    
    'Check numbers 19 and 20
    
    If thisws.Range("f29") = "1" And (thisws.Range("E52") = "" Or thisws.Range("N52") = "") Then
        MsgBox "Either number 19 and/or 20 have been left blank. Please check."
        thisws.Range("E52").Select
        End
    End If
    
    'Message that all edits passed OK
    MsgBox "All edits passed. An email will be sent to your supervisor that this review is completed."
    'send email to supervisor
    Call send_email("SNAP Negative", thisws)
    'Put cwopa and date on schedule
    Range("AQ17") = Environ("USERNAME") & " " & Date
    
End Sub
Sub snap_neg_return()
'this routine looks for problems with schedule for supervisor to return to examiner in the SNAP Neg Schedule
    Dim thisws As Worksheet
    Dim datasourcewb As Workbook, datasourcews As Worksheet
    Dim thiswb As Workbook
    
    Set thiswb = ActiveWorkbook
    Set thisws = ActiveSheet
    'send email to examiner
    Call send_email_edit("SNAP Negative", thisws)
    'Put cwopa and date on schedule
    Range("AQ20") = Environ("USERNAME") & " " & Date
     
End Sub
Sub ga_edit_check()
'this routine looks for edits in the GA Schedule
    Dim thisws As Worksheet
    Dim datasourcewb As Workbook, datasourcews As Worksheet
    Dim thiswb As Workbook
    
    Set thiswb = ActiveWorkbook
    Set thisws = ActiveSheet
    
'Check Disposition Code
If thisws.Range("AI10") = "" Then
    MsgBox "Please enter Disposition Code."
    this.Range("AI10").Select
End If
    
    
    'If Amount of Error > 0, then Review Findings must be 2 to 4.
    If thisws.Range("AO10") > 0 And Val(thisws.Range("AL10")) < 2 Then
        MsgBox "Review Findings is " & thisws.Range("AL10") & " and Amount of Error is " & thisws.Range("AO10") & _
             ".  However, if Amount of Error greater than 0, then Review Findings must be 2, 3, or 4."
        End
    End If
  
    
    'most recent action must be > or = to most recent opening
    If thisws.Range("L16") < thisws.Range("A16") Then
    MsgBox "Most recent action date (" & Format(thisws.Range("L16"), "mm/dd/yyyy") & ") must be greater than or equal to the most recent opening date (" & Format(thisws.Range("A16"), "mm/dd/yyyy") & ")."
    End
    End If
   
        

    'Message that all edits passed OK
    MsgBox "All edits passed. An email will be sent to your supervisor that this review is completed."
    'send email to supervisor
    Call send_email("GA", thisws)
    'Put cwopa and date on schedule
    Range("AU12") = Environ("USERNAME") & " " & Date
    thiswb.Worksheets("GA Workbook").Range("G38") = Date
    
End Sub
Sub tanf_edit_check()
'this routine looks for edits in the TANF Schedule
    Dim thisws As Worksheet
    Dim datasourcewb As Workbook, datasourcews As Worksheet
    Dim thiswb As Workbook
    
    Set thiswb = ActiveWorkbook
    Set thisws = ActiveSheet
    
'Check Disposition Code
If thisws.Range("AI10") = "" Then
    MsgBox "Please enter Disposition Code."
    this.Range("AI10").Select
End If
    
If thisws.Range("AI10") = "1" Then

    'check relationships in Section III
    For irow = 30 To 48 Step 2
            ' If person number is blank then end of data
            If thisws.Range("A" & irow) = "" Then
                Exit For
            End If
            'Check if Rel to Payment Name is 02, 04 or 15 to 20 and then age must be less than 20
            If (Val(thisws.Range("M" & irow)) >= 20 And (Val(thisws.Range("J" & irow)) = 2 _
                Or Val(thisws.Range("J" & irow)) = 4 Or _
               (Val(thisws.Range("J" & irow)) < 21 And Val(thisws.Range("J" & irow)) > 14))) Then
               MsgBox "In Section III, Person Number " & thisws.Range("A" & irow) & _
               " has a Relationship to Payment Name Code of 02, 04, or 15 to 20, but their age is greater than or equal to " & _
               "20 years old."
               thisws.Range("G" & irow).Select
               End
            End If
    Next irow
    
    'Set counter for number of members
    count_members = 0
    'check relationships in Section III
    For irow = 30 To 44 Step 2
        ' If person number is blank then end of data
        If thisws.Range("A" & irow) = "" Then
            Exit For
        End If
        'Sum up number of members if Dep./Rel. is between 21 and 38 or 63
        If (Val(thisws.Range("G" & irow)) > 20 And Val(thisws.Range("G" & irow)) < 39) Or _
            Val(thisws.Range("G" & irow)) = 63 Then
              count_members = count_members + 1
        End If
        ' Check if age is less than 20 and Dep./Rel. is 21, 22, 24 or 25
        If Val(thisws.Range("M" & irow)) < 20 And (Val(thisws.Range("G" & irow)) = 21 Or _
            Val(thisws.Range("G" & irow)) = 22 Or Val(thisws.Range("G" & irow)) = 24 Or _
            Val(thisws.Range("G" & irow)) = 25) Then
            'Check if Rel to Payment Name is 02, 04 or 15 to 20
            If Not (Val(thisws.Range("J" & irow)) = 2 Or Val(thisws.Range("J" & irow)) = 4 Or _
               (Val(thisws.Range("J" & irow)) < 21 And Val(thisws.Range("J" & irow)) > 14)) Then
               MsgBox "In Section III, Person Number " & thisws.Range("A" & irow) & " is less than " & _
               "20 years old and has a Dep._Rel. Code of 21,22,24 or 25 but does not have a Relationship " & _
               "to Payment Name Code of 02, 04, or 15 to 20."
               thisws.Range("G" & irow).Select
               End
            End If
        End If
         'Check to see if Educ. Level is 00 should be 98
        If thisws.Range("AI10") = 1 And Val(thisws.Range("V" & irow)) = 0 Then
            MsgBox "Education Level must be 01 - 19, 98 and 99, if Disposition = 1."
        End
        End If
    Next irow
    
    ' Check if sum of members = number of case members
    If count_members <> Val(thisws.Range("W16")) Then
        MsgBox "Count of Persons in Section III with Dep./Rel. between 21 and 38 or 63 equals " & _
        count_members & " which does not equal Number of Case Members in Section II which is " & Val(thisws.Range("W16"))
        thisws.Range("G30").Select
        End
    End If
    
    'check in Section V that income type 47 is equal to the sample month payment
    For irow = 50 To 56 Step 2
        ' If person number is blank then end of data
        If thisws.Range("C" & irow) = "" Then
            Exit For
        End If
        For icol = 7 To 37 Step 10
            'Check income type
            If Val(thisws.Cells(irow, icol)) = 47 Then
                If Val(thisws.Cells(irow, icol + 4)) <> Val(thisws.Range("I20")) Then
                    MsgBox "In Section V, Amount of income with Income Type 47 ($" & _
                        Val(thisws.Cells(irow, icol + 4)) & ") does not equal Sample Month Payment ($" & _
                        Val(thisws.Range("I20")) & ")."
                    thisws.Cells(irow, icol + 5).Select
                    End
                End If
            End If
        Next icol
    Next irow
    
    'In section V, if type of income = 11 and amount of income > 0,
    'then work activity in section III must be 01 to 04 for that person
    For irow = 50 To 56 Step 2
        ' If person number is blank then end of data
        If thisws.Range("C" & irow) = "" Then
            Exit For
        End If
        For icol = 7 To 37 Step 10
            'Check income type for 11 and amount > 0
            If Val(thisws.Cells(irow, icol)) = 11 And Val(thisws.Cells(irow, icol + 4)) > 0 Then
                'now check this person's Work Activity in Section III
                ifound = 0 'flag to check if person was found in section III
                For irowiii = 30 To 44 Step 2
                    'look for same person number from section V in section III
                    If Val(thisws.Range("C" & irow)) = Val(thisws.Range("A" & irowiii)) Then
                        ifound = 1
                        'check that work activity is between 01 to 04
                        If Val(thisws.Range("AC" & irowiii)) > 4 Then
                            MsgBox "In Section V, Person Number " & thisws.Range("C" & irow) & _
                                " has Income Type 11, but Work Activity " & thisws.Range("AC" & irowiii) & _
                                ". Work Activity should be 00 - 04 for Income Type 11."
                            thisws.Range("AC" & irowiii).Select
                            End
                        End If
                    End If
                Next irowiii
                'check to see if person from section V was found in section III
                If ifound = 0 Then
                    MsgBox "Person number " & thisws.Range("C" & irow) & " in Section V was not found in Section III."
                    End
                End If
            End If
        Next icol
    Next irow
    
    
    'check that there is a person in Section III that matches each person in Section IV
    For irow = 50 To 56 Step 2 ' section 5
        'if line is blank, then reached end of data
        If thisws.Range("C" & irow) = "" Then Exit For
        'set found flag
        found = 0
        For jrow = 30 To 44 Step 2 ' section 3
            If thisws.Range("C" & irow) = thisws.Range("A" & jrow) Then
                found = 1
            End If
        Next jrow
        If found = 0 Then 'no match was found
            thisws.Range("C" & irow).Select
            MsgBox "No match for person number " & thisws.Range("C" & irow) & " in Section IV was found in Section III."
            End
        End If
    Next irow
    
    'If Most Recent Opening = Most Recent Action date, then Action Type must be 1 or 2
'    If thisws.Range("A16") = thisws.Range("L16") And Not (Val(thisws.Range("U16")) = 1 Or Val(thisws.Range("U16")) = 2) Then
'       MsgBox "Activity Type = " & thisws.Range("U16") & " but when Most Recent Opening is the same as " & _
'        "Most Rection Action date, then Action Type must be 1 or 2."
'        End
'    End If
    
    'If Amount of Error > 0, then Review Findings must be 2 to 4.
    If thisws.Range("AO10") > 0 And Val(thisws.Range("AL10")) < 2 Then
        MsgBox "Review Findings is " & thisws.Range("AL10") & " and Amount of Error is " & thisws.Range("AO10") & _
             ".  However, if Amount of Error greater than 0, then Review Findings must be 2, 3, or 4."
        End
    End If
    
    'check if person in section V in section III
    For irow = 50 To 56 Step 2
        ' If person number is blank then end of data
        If thisws.Range("C" & irow) = "" Then
            Exit For
        End If
            For irowiii = 30 To 44 Step 2
                If Val(thisws.Range("C" & irow)) = Val(thisws.Range("A" & irowiii)) Then
                ifound = 1
                Exit For
            End If
        Next irowiii
    If found = 0 Then
        MsgBox "Person number " & thisws.Range("C" & irow) & "  in section V must be the same as person number in section III."
            End
        End If
    Next irow
                
    'sanction amount = 0 if protective payment status = 0 and disp = 1
'    If Val(thisws.Range("N20")) = 0 And Val(thisws.Range("O20")) = 0 And Val(thisws.Range("AI10")) = 1 And Val(thisws.Range("Q20")) <> 0 Then
'        MsgBox "Sanction amount must equal 0 if protective payment status = 00 and disposition = 1."
'        End
'    End If
    
    'protective payment status if sanction amount >0
'    If Val(thisws.Range("AI10")) = 1 And Val(thisws.Range("Q20")) > 0 And (Val(thisws.Range("N20")) = 0 Or Val(thisws.Range("O20")) = 0) Then
'        MsgBox "Protective payment status must be 11-16, 21-26, 31-36, 41-46, 56 if sanction amount is > 0 and dispoition =1."
'        End
'    End If
    
    'GCI, ID, NCI, FSA, OPR, and TANF Days can't be blank
    If thisws.Range("AI10") = 1 Then
        If thisws.Range("AB20") = "" Or thisws.Range("AH20") = "" Or thisws.Range("AN20") = "" Or thisws.Range("B24") = "" Or thisws.Range("U24") = "" Or thisws.Range("AN24") = "" Then
        MsgBox "Gross Countable Income, Income Disregard, Net Countable Income, Food Stamp Allotment, Overpayment Recoupment and TANF Days cannot be blank if Disposition = 1."
            End
        End If
    End If
    
    'most recent action must be > or = to most recent opening
    If thisws.Range("L16") < thisws.Range("A16") Then
    MsgBox "Most recent action date (" & Format(thisws.Range("L16"), "mm/dd/yyyy") & ") must be greater than or equal to the most recent opening date (" & Format(thisws.Range("A16"), "mm/dd/yyyy") & ")."
    End
    End If
    
    'most recent opening or most recent action must be before the end of the review month
    'calculate end of review month - increment the month by 1 and set date to 0
    reviewmonthend = DateSerial(Val(Right(thisws.Range("AB10"), 4)), Val(Left(thisws.Range("AB10"), 2)) + 1, 0)
    'check most recent opening date
    If thisws.Range("A16") > reviewmonthend Then
        MsgBox "Most Recent Opening date (" & Format(thisws.Range("A16"), "mm/dd/yyyy") & ") is after the end of the review month (" & _
            Format(reviewmonthend, "mm/dd/yyyy") & ")."
        End
    End If
    'check most recent action date
    If thisws.Range("L16") > reviewmonthend Then
        MsgBox "Most Recent Action date (" & Format(thisws.Range("L16"), "mm/dd/yyyy") & ") is after the end of the review month (" & _
            Format(reviewmonthend, "mm/dd/yyyy") & ")."
        End
    End If
        
End If

    'Message that all edits passed OK
    MsgBox "All edits passed. An email will be sent to your supervisor that this review is completed."
    'send email to supervisor
    Call send_email("TANF", thisws)
    'Put cwopa and date on schedule
    Range("AU12") = Environ("USERNAME") & " " & Date
    thiswb.Worksheets("TANF Workbook").Range("G38") = Date
    
End Sub
Sub tanf_return()
'this routine looks for problems with schedule for supervisor to return to examiner in the TANF Schedule
    Dim thisws As Worksheet
    Dim datasourcewb As Workbook, datasourcews As Worksheet
    Dim thiswb As Workbook
    
    Set thiswb = ActiveWorkbook
    Set thisws = ActiveSheet
    'send email to examiner
    Call send_email_edit("TANF", thisws)
    'Put cwopa and date on schedule
    Range("BV3") = Environ("USERNAME") & " " & Date
     
End Sub

Sub ma_edit_check_pos()
'this routine looks for edits in the MA Positive Schedule
    Dim thisws As Worksheet
    Dim datasourcewb As Workbook, datasourcews As Worksheet
    Dim thiswb As Workbook
    
    Set thiswb = ActiveWorkbook
    Set thisws = ActiveSheet
    
    'check Person Number in Section III for single digits and make them double digit
    For irow = 51 To 73 Step 2
        If Len(thisws.Range("B" & irow)) = 1 Then
            thisws.Range("B" & irow) = "0" & thisws.Range("B" & irow)
        End If
    Next irow
    
    'Check if COVID is blank
    If thisws.Range("AK10") = "" And thisws.Range("F16") = 1 Then
        MsgBox "Please enter a COVID Code."
        thisws.Range("AK10").Select
        End
    End If
    
     'Check if prior assistance code is blank
    If thisws.Range("J27") = "" And thisws.Range("F16") = 1 Then
        MsgBox "Please enter a Prior Assistance Code."
        thisws.Range("J27").Select
        End
    End If
    
    'Check if action code is blank
    If thisws.Range("V27") = "" And thisws.Range("F16") = 1 Then
        MsgBox "Please enter an Action Code."
        thisws.Range("V27").Select
        End
    End If
    
    'check Person Number in Section IV for single digits and make them double digit
    For irow = 78 To 84 Step 2
        If Len(thisws.Range("B" & irow)) = 1 Then
            thisws.Range("B" & irow) = "0" & thisws.Range("B" & irow)
        End If
    Next irow
    
    'check that there is a person in Section III that matches each person in Section IV
    For irow = 78 To 84 Step 2
        'if line is blank, then reached end of data
        If thisws.Range("B" & irow) = "" Then Exit For
        'set found flag
        found = 0
        For jrow = 51 To 73 Step 2
            If thisws.Range("B" & irow) = thisws.Range("B" & jrow) Then
                found = 1
            End If
        Next jrow
        If found = 0 Then 'no match was found
            thisws.Range("B" & irow).Select
            MsgBox "No match for person number " & thisws.Range("B" & irow) & " in Section IV was found in Section III."
            End
        End If
    Next irow
    
    'Check Gross Countable Income
    sum_income = 0
    For irow = 78 To 84 Step 2
        For icol = 10 To 40 Step 10
            sum_income = sum_income + thisws.Cells(irow, icol)
        Next icol
    Next irow
    If sum_income <> thisws.Range("AF32") Then
        thisws.Range("B78").Select
        MsgBox "Sum of incomes ($" & sum_income & ") in section IV did not match Gross Countable Income ($" & thisws.Range("AF32") & ")."
        End
    End If
    
    'Check age and Employment Status
    For irow = 51 To 73 Step 2
        If thisws.Range("R" & irow) < 16 And thisws.Range("R" & irow) <> "" Then
            thisws.Range("AM" & irow) = "-"
        End If
    Next irow
    
    'Check dates in Section V
    For irow = 96 To 112 Step 2
        If thisws.Range("AL" & irow) = "" Then Exit For
        sample_date = DateSerial(Right(thisws.Range("AM10"), 4), Left(thisws.Range("AM10"), 2), 1)
        If thisws.Range("AL" & irow) > sample_date Then
            thisws.Range("AL" & irow).Select
            MsgBox "One of the dates " & thisws.Range("AL" & irow) & " in Section V is later than the Sample Month."
            End
        End If
    Next irow
    
    'Message that all edits passed OK
    MsgBox "All edits passed. An email will be sent to your supervisor that this review is completed."
    'send email to supervisor
    Call send_email("MA Positive", thisws)
    'Put cwopa and date on schedule
    Range("CD13") = Environ("USERNAME") & " " & Date
    thiswb.Worksheets("MA Workbook").Range("F41") = Date
    
End Sub
Sub MA_return()
'this routine looks for problems with schedule for supervisor to return to examiner in the MA Positive Schedule
    Dim thisws As Worksheet
    Dim datasourcewb As Workbook, datasourcews As Worksheet
    Dim thiswb As Workbook
    
    Set thiswb = ActiveWorkbook
    Set thisws = ActiveSheet
    'send email to examiner
    Call send_email_edit("MA Positive", thisws)
    'Put cwopa and date on schedule
    Range("CK8") = Environ("USERNAME") & " " & Date
     
End Sub

Sub send_email(review_type As String, thisws As Worksheet)

'send email to supervisor

    Dim DLetter As String, DUNC As String
    Dim datasourcewb As Workbook, datasourcews As Worksheet
    Dim sPath As String, PathStr As String
    Dim OutApp As Object, OutMail As Object
    
    Application.ScreenUpdating = False

    'Find path to Examiner's files on this PC
    Set WshNetwork = CreateObject("WScript.Network")
    Set oDrives = WshNetwork.EnumNetworkDrives

    DLetter = ""
    For i = 0 To oDrives.Count - 1 Step 2
        DUNC = "" & oDrives.Item(i + 1) & ""
    If LCase(DUNC) = "\\hsedcprapfpp001\oim\pwimdaubts04\data\stat" Then
        DLetter = "" & oDrives.Item(i) & "\DQC\"
        Exit For
    ElseIf LCase(DUNC) = "\\hsedcprapfpp001\oim\pwimdaubts04\data\stat\dqc" Then
        DLetter = "" & oDrives.Item(i) & "\"
        Exit For
    ElseIf LCase(DUNC) = "\\hsedcprapfpp001\oim\pwimdaubts04\data\stat" Then
        DLetter = "" & oDrives.Item(i) & "\DQC"
        Exit For
    End If

    Next i

    If DLetter = "" Then
        MsgBox "Network Drive to DQC Directory is NOT correct" & Chr(13) & _
            "Contact Nicole or Valerie"
        End
    End If

    PathStr = DLetter    'Copy template data source to active directory
    sPath = ActiveWorkbook.Path
    DestFile = sPath & "\FM DS Temp.xlsx"
    SrceFile = PathStr & "Finding Memo\Finding Memo Data Source.xlsx"
    FileCopy SrceFile, DestFile

    'Open spreadsheet where all the data is stored to populate findings memo
    Workbooks.Open fileName:=DestFile, UpdateLinks:=False

    Set datasourcewb = ActiveWorkbook
    Set datasourcews = ActiveSheet

    Set OutApp = CreateObject("Outlook.Application")
    OutApp.Session.Logon
    Set OutMail = OutApp.CreateItem(0)
    
    'check if review type to find review number, month and reviewer number on schedule
    Select Case review_type
        Case "MA Positive"
            'review number
            review_number = thisws.Range("A10") & thisws.Range("B10") & thisws.Range("c10") & thisws.Range("D10") & thisws.Range("E10") & thisws.Range("F10")
            'review month
            sample_month = thisws.Range("AM10")
            'reviewer number
            datasourcews.Range("A7") = Val(thisws.Range("AO3") & thisws.Range("AP3"))
            'Drop/Finding
            If Val(thisws.Range("F16")) = 1 Then
                If Val(thisws.Range("S16")) = 1 Then
                    finding_code = " clean."
                Else
                    finding_code = " in error."
                End If
            Else
                finding_code = " a drop."
            End If
        Case "MA Negative"
            'review number
            review_number = thisws.Range("L15")
            'review month
            sample_month = thisws.Range("T11")
            'reviewer number
            datasourcews.Range("A7") = Val(thisws.Range("AB11") & thisws.Range("AC11"))
            'Drop/Finding
            If Val(thisws.Range("M56")) = 0 Then
                If Val(thisws.Range("C40")) = 1 And (Val(thisws.Range("C25")) = 1 Or Val(thisws.Range("C25")) = 5) Then
                    finding_code = " clean."
                Else
                    finding_code = " in error."
                End If
            Else
                finding_code = " a drop."
            End If
        Case "TANF"
            'review number
            review_number = thisws.Range("A10")
            'review month
            sample_month = thisws.Range("AB10")
            'reviewer number
            datasourcews.Range("A7") = Val(thisws.Range("AO3") & thisws.Range("AP3"))
            'Drop/Finding
            If Val(thisws.Range("AI10")) = 1 Then
                If Val(thisws.Range("AL10")) = 1 Then
                    finding_code = " clean."
                Else
                    finding_code = " in error."
                End If
            Else
                finding_code = " a drop."
            End If
         Case "GA"
            'review number
            review_number = thisws.Range("A10")
            'review month
            sample_month = thisws.Range("AB10")
            'reviewer number
            datasourcews.Range("A7") = Val(thisws.Range("AO3") & thisws.Range("AP3"))
            'Drop/Finding
            If Val(thisws.Range("AI10")) = 1 Then
                If Val(thisws.Range("AL10")) = 1 Then
                    finding_code = " clean."
                Else
                    finding_code = " in error."
                End If
            Else
                finding_code = " a drop."
            End If
        Case "SNAP Positive"
            'review number
            review_number = thisws.Range("A18")
            'review month
            sample_month = thisws.Range("AD18") & thisws.Range("AG18")
            'reviewer number
            datasourcews.Range("A7") = Val(thisws.Range("AJ5") & thisws.Range("AK5"))
            'Drop/Finding
            If Val(thisws.Range("C22")) = 1 Then
                If Val(thisws.Range("K22")) = 1 Then
                    finding_code = " clean."
                Else
                    finding_code = " in error."
                End If
            Else
                finding_code = " a drop."
            End If
        Case "SNAP Negative"
            'review number
            review_number = thisws.Range("C20")
            'review month
            sample_month = thisws.Range("AF20") & thisws.Range("AI20")
            'reviewer number
            datasourcews.Range("A7") = Val(thisws.Range("W17"))
            'Finding
            'If Val(thisws.Range("AF34")) = 1 Then 'old code
            If Val(thisws.Range("F29")) = 1 Then 'new code 11/2014
                'If Val(thisws.Range("AF29")) = 1 Then 'old code
                If Val(thisws.Range("M29")) = 1 Then 'new code 11/2014
                    finding_code = " clean."
                Else
                    finding_code = " in error."
                End If
            Else
                finding_code = " a drop."
            End If
    End Select
    
    'set up email
    With OutMail
        .To = datasourcews.Range("AO4")
        '.To = "wcohick@pa.gov"
        .Subject = "Completion of " & review_type & " Schedule Review " & review_number & " Month " & _
            sample_month & " Examiner " & datasourcews.Range("A7")
    
        .HTMLBody = "Hi " & datasourcews.Range("AP2") & ", " & "<br><br>" & datasourcews.Range("AM2") & _
              " has completed the " & review_type & " Schedule Review " & _
              review_number & " for Sample Month " & sample_month & ". This case is" & finding_code
        '.DeleteAfterSubmit = True
        .NoAging = True
        .Display  'Or use Display
    End With
    
    'send Email
    Set OutMail = Nothing
    SendKeys "%{s}", True 'send the email without prompts
    
   Set OutApp = Nothing
   
   datasourcewb.Close False
   Kill DestFile
    Application.ScreenUpdating = True

End Sub
Sub send_email_edit(review_type As String, thisws As Worksheet)

'send email to examiner

    Dim DLetter As String, DUNC As String
    Dim datasourcewb As Workbook, datasourcews As Worksheet
    Dim sPath As String, PathStr As String
    Dim OutApp As Object, OutMail As Object
    Dim myValue As Variant
    
    Application.ScreenUpdating = False

    'Find path to Examiner's files on this PC
    Set WshNetwork = CreateObject("WScript.Network")
    Set oDrives = WshNetwork.EnumNetworkDrives

    DLetter = ""
    For i = 0 To oDrives.Count - 1 Step 2
         DUNC = "" & oDrives.Item(i + 1) & ""
    If LCase(DUNC) = "\\hsedcprapfpp001\oim\pwimdaubts04\data\stat" Then
        DLetter = "" & oDrives.Item(i) & "\DQC\"
        Exit For
    ElseIf LCase(DUNC) = "\\hsedcprapfpp001\oim\pwimdaubts04\data\stat\dqc" Then
        DLetter = "" & oDrives.Item(i) & "\"
        Exit For
    ElseIf LCase(DUNC) = "\\hsedcprapfpp001\oim\pwimdaubts04\data\stat" Then
        DLetter = "" & oDrives.Item(i) & "\DQC"
        Exit For
    End If

    Next i

    If DLetter = "" Then
        MsgBox "Network Drive to DQC Directory is NOT correct" & Chr(13) & _
            "Contact Nicole or Valerie"
        End
    End If

    PathStr = DLetter    'Copy template data source to active directory
    sPath = ActiveWorkbook.Path
    DestFile = sPath & "\FM DS Temp.xlsx"
    SrceFile = PathStr & "Finding Memo\Finding Memo Data Source.xlsx"
    FileCopy SrceFile, DestFile

    'Open spreadsheet where all the data is stored to populate findings memo
    Workbooks.Open fileName:=DestFile, UpdateLinks:=False

    Set datasourcewb = ActiveWorkbook
    Set datasourcews = ActiveSheet

    Set OutApp = CreateObject("Outlook.Application")
    OutApp.Session.Logon
    Set OutMail = OutApp.CreateItem(0)
    
    'check if review type to find review number, month and reviewer number on schedule
    Select Case review_type
        Case "MA Positive"
            'review number
            review_number = thisws.Range("A10") & thisws.Range("B10") & thisws.Range("c10") & thisws.Range("D10") & thisws.Range("E10") & thisws.Range("F10")
            'review month
            sample_month = thisws.Range("AM10")
            'reviewer number
            datasourcews.Range("A7") = Val(thisws.Range("AO3") & thisws.Range("AP3"))
        Case "MA Negative"
            'review number
            review_number = thisws.Range("L15")
            'review month
            sample_month = thisws.Range("T11")
            'reviewer number
            datasourcews.Range("A7") = Val(thisws.Range("AB11") & thisws.Range("AC11"))
        Case "TANF"
            'review number
            review_number = thisws.Range("A10")
            'review month
            sample_month = thisws.Range("AB10")
            'reviewer number
            datasourcews.Range("A7") = Val(thisws.Range("AO3") & thisws.Range("AP3"))
        Case "SNAP Positive"
            'review number
            review_number = thisws.Range("A18")
            'review month
            sample_month = thisws.Range("AD18") & thisws.Range("AG18")
            'reviewer number
            datasourcews.Range("A7") = Val(thisws.Range("AJ5") & thisws.Range("AK5"))
        Case "SNAP Negative"
            'review number
            review_number = thisws.Range("C20")
            'review month
            sample_month = thisws.Range("AF20") & thisws.Range("AI20")
            'reviewer number
            datasourcews.Range("A7") = Val(thisws.Range("W17"))
    End Select
    
    'set up email
    With OutMail
        .To = datasourcews.Range("AM4")
        '.CC = datasourcews.Range("AP2")
        
       ' .To = "vamiller@pa.gov"
        .Subject = "Return to Examiner " & review_type & " Schedule Review " & review_number & " Month " & _
            sample_month & " Examiner " & datasourcews.Range("A7")

      .HTMLBody = "Hi " & datasourcews.Range("AM2") & ", " & "<br><br>" & datasourcews.Range("AP2") & _
              " found the following error in your review " & InputBox("Why are you returning review to examiner?") & " for review number " & _
              review_number & " for Sample Month " & sample_month & "."
        '.DeleteAfterSubmit = True
        .NoAging = True
        .Display  'Or use Display
    End With
    
    'send Email
    Set OutMail = Nothing
    SendKeys "%{s}", False 'send the email without prompts
    
   Set OutApp = Nothing
   
   datasourcewb.Close False
   Kill DestFile
   Application.ScreenUpdating = True

End Sub

Sub ma_edit_check_neg()
'this routine looks for edits in the MA Negative Schedule
    Dim thisws As Worksheet
    Dim datasourcewb As Workbook, datasourcews As Worksheet
    
    Set thisws = ActiveSheet
    
    'If drop, then put blanks in 5 cells
    If thisws.Range("M56") <> 0 Then
        thisws.Range("J19") = ""
        thisws.Range("S19") = ""
        thisws.Range("C25") = ""
        thisws.Range("Q25") = ""
        thisws.Range("C40") = ""
    End If
    
    'make sure that box J is blank
    thisws.Range("C56") = ""
    
    'Message that all edits passed OK
    MsgBox "All edits passed. An email will be sent to your supervisor that this review is completed."
    'send email
    Call send_email("MA Negative", thisws)
    'Put cwopa and date on schedule
    Range("AI9") = Environ("USERNAME") & " " & Date
    Range("L11") = Date
    
End Sub
Sub MA_neg_return()
'this routine looks for problems with schedule for supervisor to return to examiner in the MA Negative Schedule
    Dim thisws As Worksheet
    Dim datasourcewb As Workbook, datasourcews As Worksheet
    Dim thiswb As Workbook
    
    Set thiswb = ActiveWorkbook
    Set thisws = ActiveSheet
    'send email to examiner
    Call send_email_edit("MA Negative", thisws)
    'Put cwopa and date on schedule
    Range("AO3") = Environ("USERNAME") & " " & Date
     
End Sub

Sub OB24_Click()
    If ActiveSheet.Shapes("OB 24").OLEFormat.Object.Value = 1 Then
        ActiveSheet.Shapes("OB 2").Visible = msoFalse
        ActiveSheet.Shapes("OB 3").Visible = msoFalse
        ActiveSheet.Shapes("OB 4").Visible = msoFalse
        ActiveSheet.Shapes("OB 5").Visible = msoFalse
        ActiveSheet.Shapes("OB 6").Visible = msoFalse
        ActiveSheet.Shapes("OB 121").Visible = msoFalse
        ActiveSheet.Shapes("OB 119").Visible = msoFalse
        ActiveSheet.Shapes("OB 120").Visible = msoFalse
        ActiveSheet.Shapes("OB 7").Visible = msoFalse
        ActiveSheet.Shapes("OB 8").Visible = msoFalse
        ActiveSheet.Shapes("OB 65").Visible = msoFalse
        ActiveSheet.Shapes("OB 66").Visible = msoFalse
        ActiveSheet.Shapes("OB 70").Visible = msoFalse
        ActiveSheet.Shapes("OB 72").Visible = msoFalse
        ActiveSheet.Shapes("OB 73").Visible = msoFalse
        ActiveSheet.Shapes("OB 78").Visible = msoFalse
        ActiveSheet.Shapes("OB 79").Visible = msoFalse
        ActiveSheet.Shapes("OB 83").Visible = msoFalse
        ActiveSheet.Shapes("OB 86").Visible = msoFalse
        ActiveSheet.Shapes("OB 87").Visible = msoFalse
        ActiveSheet.Shapes("OB 11").Visible = msoFalse
        ActiveSheet.Shapes("OB 22").Visible = msoFalse
        ActiveSheet.Shapes("OB 92").Visible = msoFalse
        ActiveSheet.Shapes("OB 93").Visible = msoFalse
        ActiveSheet.Shapes("OB 95").Visible = msoFalse
        ActiveSheet.Shapes("OB 97").Visible = msoFalse
        ActiveSheet.Shapes("OB 96").Visible = msoFalse
        ActiveSheet.Shapes("OB 12").Visible = msoFalse
        ActiveSheet.Shapes("OB 13").Visible = msoFalse
        ActiveSheet.Shapes("OB 61").Visible = msoFalse
        ActiveSheet.Shapes("OB 14").Visible = msoFalse
        ActiveSheet.Shapes("OB 15").Visible = msoFalse
        ActiveSheet.Shapes("OB 117").Visible = msoFalse
        ActiveSheet.Shapes("OB 62").Visible = msoFalse
        ActiveSheet.Shapes("OB 16").Visible = msoFalse
        ActiveSheet.Shapes("OB 17").Visible = msoFalse
        ActiveSheet.Shapes("OB 107").Visible = msoFalse
        ActiveSheet.Shapes("OB 108").Visible = msoFalse
        ActiveSheet.Shapes("OB 111").Visible = msoFalse
        ActiveSheet.Shapes("OB 110").Visible = msoFalse
        ActiveSheet.Shapes("OB 112").Visible = msoFalse
        'ActiveSheet.Shapes("OB 9").Visible = msoFalse
        'ActiveSheet.Shapes("OB 10").Visible = msoFalse
        'ActiveSheet.Shapes("OB 88").Visible = msoFalse
        'ActiveSheet.Shapes("OB 89").Visible = msoFalse
        'ActiveSheet.Shapes("OB 103").Visible = msoFalse
        'ActiveSheet.Shapes("OB 104").Visible = msoFalse
    End If
End Sub

Sub OB23_Click()
    If ActiveSheet.Shapes("OB 23").OLEFormat.Object.Value = 1 Then
        ActiveSheet.Shapes("OB 2").Visible = msoTrue
        ActiveSheet.Shapes("OB 2").Visible = msoTrue
        ActiveSheet.Shapes("OB 3").Visible = msoTrue
        ActiveSheet.Shapes("OB 4").Visible = msoTrue
        ActiveSheet.Shapes("OB 5").Visible = msoTrue
        ActiveSheet.Shapes("OB 6").Visible = msoTrue
        ActiveSheet.Shapes("OB 121").Visible = msoTrue
        ActiveSheet.Shapes("OB 119").Visible = msoTrue
        ActiveSheet.Shapes("OB 120").Visible = msoTrue
        ActiveSheet.Shapes("OB 7").Visible = msoTrue
        ActiveSheet.Shapes("OB 8").Visible = msoTrue
        ActiveSheet.Shapes("OB 65").Visible = msoTrue
        ActiveSheet.Shapes("OB 66").Visible = msoTrue
        ActiveSheet.Shapes("OB 70").Visible = msoTrue
        ActiveSheet.Shapes("OB 72").Visible = msoTrue
        ActiveSheet.Shapes("OB 73").Visible = msoTrue
        ActiveSheet.Shapes("OB 78").Visible = msoTrue
        ActiveSheet.Shapes("OB 79").Visible = msoTrue
        ActiveSheet.Shapes("OB 83").Visible = msoTrue
        ActiveSheet.Shapes("OB 86").Visible = msoTrue
        ActiveSheet.Shapes("OB 87").Visible = msoTrue
        ActiveSheet.Shapes("OB 11").Visible = msoTrue
        ActiveSheet.Shapes("OB 22").Visible = msoTrue
        ActiveSheet.Shapes("OB 92").Visible = msoTrue
        ActiveSheet.Shapes("OB 93").Visible = msoTrue
        ActiveSheet.Shapes("OB 95").Visible = msoTrue
        ActiveSheet.Shapes("OB 97").Visible = msoTrue
        ActiveSheet.Shapes("OB 96").Visible = msoTrue
        ActiveSheet.Shapes("OB 12").Visible = msoTrue
        ActiveSheet.Shapes("OB 13").Visible = msoTrue
        ActiveSheet.Shapes("OB 61").Visible = msoTrue
        ActiveSheet.Shapes("OB 14").Visible = msoTrue
        ActiveSheet.Shapes("OB 15").Visible = msoTrue
        ActiveSheet.Shapes("OB 117").Visible = msoTrue
        ActiveSheet.Shapes("OB 62").Visible = msoTrue
        ActiveSheet.Shapes("OB 16").Visible = msoTrue
        ActiveSheet.Shapes("OB 17").Visible = msoTrue
        ActiveSheet.Shapes("OB 107").Visible = msoTrue
        ActiveSheet.Shapes("OB 108").Visible = msoTrue
        ActiveSheet.Shapes("OB 111").Visible = msoTrue
        ActiveSheet.Shapes("OB 110").Visible = msoTrue
        ActiveSheet.Shapes("OB 112").Visible = msoTrue
        'ActiveSheet.Shapes("OB 9").Visible = msoTrue
        'ActiveSheet.Shapes("OB 10").Visible = msoTrue
        'ActiveSheet.Shapes("OB 88").Visible = msoTrue
        'ActiveSheet.Shapes("OB 89").Visible = msoTrue
        'ActiveSheet.Shapes("OB 103").Visible = msoTrue
        'ActiveSheet.Shapes("OB 104").Visible = msoTrue
    End If
End Sub

Sub OB131_Click()
    If ActiveSheet.Shapes("OB 131").OLEFormat.Object.Value = 1 Then
        ActiveSheet.Shapes("OB 23").Visible = msoFalse
        ActiveSheet.Shapes("OB 24").Visible = msoFalse
        ActiveSheet.Shapes("OB 2").Visible = msoFalse
        ActiveSheet.Shapes("OB 3").Visible = msoFalse
        ActiveSheet.Shapes("OB 4").Visible = msoFalse
        ActiveSheet.Shapes("OB 5").Visible = msoFalse
        ActiveSheet.Shapes("OB 6").Visible = msoFalse
        ActiveSheet.Shapes("OB 121").Visible = msoFalse
        ActiveSheet.Shapes("OB 119").Visible = msoFalse
        ActiveSheet.Shapes("OB 120").Visible = msoFalse
        ActiveSheet.Shapes("OB 7").Visible = msoFalse
        ActiveSheet.Shapes("OB 8").Visible = msoFalse
        ActiveSheet.Shapes("OB 65").Visible = msoFalse
        ActiveSheet.Shapes("OB 66").Visible = msoFalse
        ActiveSheet.Shapes("OB 70").Visible = msoFalse
        ActiveSheet.Shapes("OB 72").Visible = msoFalse
        ActiveSheet.Shapes("OB 73").Visible = msoFalse
        ActiveSheet.Shapes("OB 78").Visible = msoFalse
        ActiveSheet.Shapes("OB 79").Visible = msoFalse
        ActiveSheet.Shapes("OB 83").Visible = msoFalse
        ActiveSheet.Shapes("OB 86").Visible = msoFalse
        ActiveSheet.Shapes("OB 87").Visible = msoFalse
        ActiveSheet.Shapes("OB 11").Visible = msoFalse
        ActiveSheet.Shapes("OB 22").Visible = msoFalse
        ActiveSheet.Shapes("OB 92").Visible = msoFalse
        ActiveSheet.Shapes("OB 93").Visible = msoFalse
        ActiveSheet.Shapes("OB 95").Visible = msoFalse
        ActiveSheet.Shapes("OB 97").Visible = msoFalse
        ActiveSheet.Shapes("OB 96").Visible = msoFalse
        ActiveSheet.Shapes("OB 12").Visible = msoFalse
        ActiveSheet.Shapes("OB 13").Visible = msoFalse
        ActiveSheet.Shapes("OB 61").Visible = msoFalse
        ActiveSheet.Shapes("OB 14").Visible = msoFalse
        ActiveSheet.Shapes("OB 15").Visible = msoFalse
        ActiveSheet.Shapes("OB 117").Visible = msoFalse
        ActiveSheet.Shapes("OB 62").Visible = msoFalse
        ActiveSheet.Shapes("OB 16").Visible = msoFalse
        ActiveSheet.Shapes("OB 17").Visible = msoFalse
        ActiveSheet.Shapes("OB 107").Visible = msoFalse
        ActiveSheet.Shapes("OB 108").Visible = msoFalse
        ActiveSheet.Shapes("OB 111").Visible = msoFalse
        ActiveSheet.Shapes("OB 110").Visible = msoFalse
        ActiveSheet.Shapes("OB 112").Visible = msoFalse
        'ActiveSheet.Shapes("OB 9").Visible = msoFalse
        'ActiveSheet.Shapes("OB 10").Visible = msoFalse
        'ActiveSheet.Shapes("OB 88").Visible = msoFalse
        'ActiveSheet.Shapes("OB 89").Visible = msoFalse
        'ActiveSheet.Shapes("OB 103").Visible = msoFalse
        'ActiveSheet.Shapes("OB 104").Visible = msoFalse
    Else
        ActiveSheet.Shapes("OB 23").Visible = msoTrue
        ActiveSheet.Shapes("OB 24").Visible = msoTrue
        ActiveSheet.Shapes("OB 2").Visible = msoTrue
        ActiveSheet.Shapes("OB 3").Visible = msoTrue
        ActiveSheet.Shapes("OB 4").Visible = msoTrue
        ActiveSheet.Shapes("OB 5").Visible = msoTrue
        ActiveSheet.Shapes("OB 6").Visible = msoTrue
        ActiveSheet.Shapes("OB 121").Visible = msoTrue
        ActiveSheet.Shapes("OB 119").Visible = msoTrue
        ActiveSheet.Shapes("OB 120").Visible = msoTrue
        ActiveSheet.Shapes("OB 7").Visible = msoTrue
        ActiveSheet.Shapes("OB 8").Visible = msoTrue
        ActiveSheet.Shapes("OB 65").Visible = msoTrue
        ActiveSheet.Shapes("OB 66").Visible = msoTrue
        ActiveSheet.Shapes("OB 70").Visible = msoTrue
        ActiveSheet.Shapes("OB 72").Visible = msoTrue
        ActiveSheet.Shapes("OB 73").Visible = msoTrue
        ActiveSheet.Shapes("OB 78").Visible = msoTrue
        ActiveSheet.Shapes("OB 79").Visible = msoTrue
        ActiveSheet.Shapes("OB 83").Visible = msoTrue
        ActiveSheet.Shapes("OB 86").Visible = msoTrue
        ActiveSheet.Shapes("OB 87").Visible = msoTrue
        ActiveSheet.Shapes("OB 11").Visible = msoTrue
        ActiveSheet.Shapes("OB 22").Visible = msoTrue
        ActiveSheet.Shapes("OB 92").Visible = msoTrue
        ActiveSheet.Shapes("OB 93").Visible = msoTrue
        ActiveSheet.Shapes("OB 95").Visible = msoTrue
        ActiveSheet.Shapes("OB 97").Visible = msoTrue
        ActiveSheet.Shapes("OB 96").Visible = msoTrue
        ActiveSheet.Shapes("OB 12").Visible = msoTrue
        ActiveSheet.Shapes("OB 13").Visible = msoTrue
        ActiveSheet.Shapes("OB 61").Visible = msoTrue
        ActiveSheet.Shapes("OB 14").Visible = msoTrue
        ActiveSheet.Shapes("OB 15").Visible = msoTrue
        ActiveSheet.Shapes("OB 117").Visible = msoTrue
        ActiveSheet.Shapes("OB 62").Visible = msoTrue
        ActiveSheet.Shapes("OB 16").Visible = msoTrue
        ActiveSheet.Shapes("OB 17").Visible = msoTrue
        ActiveSheet.Shapes("OB 107").Visible = msoTrue
        ActiveSheet.Shapes("OB 108").Visible = msoTrue
        ActiveSheet.Shapes("OB 111").Visible = msoTrue
        ActiveSheet.Shapes("OB 110").Visible = msoTrue
        ActiveSheet.Shapes("OB 112").Visible = msoTrue
        'ActiveSheet.Shapes("OB 9").Visible = msoTrue
        'ActiveSheet.Shapes("OB 10").Visible = msoTrue
        'ActiveSheet.Shapes("OB 88").Visible = msoTrue
        'ActiveSheet.Shapes("OB 89").Visible = msoTrue
        'ActiveSheet.Shapes("OB 103").Visible = msoTrue
        'ActiveSheet.Shapes("OB 104").Visible = msoTrue
    End If
End Sub






