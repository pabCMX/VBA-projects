Attribute VB_Name = "Populate_MA_delimited_mod"
Sub Populate_MA_Delimited()

Dim wb_sch As Workbook, review_number As String, case_row As Range

'income_freq = Array(0, 1, 4, 2, 2, 1, 0.5, 0.333333, 0.166667, 0.083333)

Set wb_sch = wb
    For Each ws In wb_sch.Worksheets
        If Val(ws.Name) > 1000 Then
            review_number = ws.Name
            Exit For
        End If
    Next
        
    'Find max row in bis case record
        maxrow_bis_case = wb_bis.Worksheets("Case").Cells.Find(What:="*", _
            SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
    'Find max column in bis case record
        LastColumn_bis_case = wb_bis.Worksheets("Case").Cells.Find(What:="*", After:=[A1], _
            SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    
    'Find what row in the bis case file the review number is on
    With wb_bis.Worksheets("Case").Range("C2:C" & maxrow_bis_case)
        'temp = Val(review_number)
        Set case_row = .Find(Val(review_number), LookIn:=xlValues)
    End With
    
    'Find max row in bis individual record
        maxrow_bis_ind = wb_bis.Worksheets("Individual").Cells.Find(What:="*", _
        SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
    'Find max column in individual case record
        LastColumn_bis_ind = wb_bis.Worksheets("Individual").Cells.Find(What:="*", After:=[A1], _
            SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    
    'Find start and stop row of the review number in individual file
    bis_ind_start_row = 0
    For i = 2 To maxrow_bis_ind
        If Val(wb_bis.Worksheets("Individual").Range("C" & i)) = review_number And bis_ind_start_row = 0 Then
            bis_ind_start_row = i
            If i = maxrow_bis_ind Then
                bis_ind_stop_row = i
                Exit For
            End If
        ElseIf Val(wb_bis.Worksheets("Individual").Range("C" & i)) <> review_number And bis_ind_start_row <> 0 Then
            bis_ind_stop_row = i - 1
            Exit For
        End If
    Next
   
    'Sort the individuals so head of household is at top
'    If bis_ind_stop_row > bis_ind_start_row Then
'        'wb_bis.Worksheets("Individual").Range(Cells(bis_ind_start_row, 1), Cells(bis_ind_stop_row, LastColumn_bis_ind)).Sort key1:=wb_bis.Worksheets("Individual").Range("T" & bis_ind_start_row & ":T" & bis_ind_stop_row), order1:=xlDescending, Header:=xlNo
'        'wb_bis.Worksheets("Individual").Range(Cells(bis_ind_start_row + 1, 1), Cells(bis_ind_stop_row, LastColumn_bis_ind)).Sort key1:=Range("L" & bis_ind_start_row + 1 & ":L" & bis_ind_stop_row), order1:=xlDescending, Header:=xlNo
'
'        wb_bis.Worksheets("Individual").Sort.SortFields.Clear
'        wb_bis.Worksheets("Individual").Sort.SortFields.Add Key:=Range( _
'            "T" & bis_ind_start_row & ":T" & bis_ind_stop_row), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
'            xlSortNormal
'        With wb_bis.Worksheets("Individual").Sort
'        .SetRange Range(Cells(bis_ind_start_row, 1), Cells(bis_ind_stop_row, LastColumn_bis_ind))
'        .Header = xlNo
'        .MatchCase = False
'        .Orientation = xlTopToBottom
'        .SortMethod = xlPinYin
'        .Apply
'        End With
'
'        wb_bis.Worksheets("Individual").Sort.SortFields.Clear
'        wb_bis.Worksheets("Individual").Sort.SortFields.Add Key:=Range( _
'            "L" & bis_ind_start_row + 1 & ":L" & bis_ind_stop_row), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
'            xlSortNormal
'        With wb_bis.Worksheets("Individual").Sort
'        .SetRange Range(Cells(bis_ind_start_row + 1, 1), Cells(bis_ind_stop_row, LastColumn_bis_ind))
'        .Header = xlNo
'        .MatchCase = False
'        .Orientation = xlTopToBottom
'        .SortMethod = xlPinYin
'        .Apply
'        End With
'    End If
    
    'Put in worksheet data from case record
    Select Case program
    
        Case "MA Positive"
            
            wb_sch.Worksheets("MA Workbook").Range("D20") = _
                wb_bis.Worksheets("Case").Range("AB" & case_row.Row) 'Telephone Number
            
            time_zero = DateSerial(0, 0, 1) 'Time format with all zeros except day = 1
                     
            'get dates from delimited file
            open_date = DateSerial(Val(wb_bis.Worksheets("Case").Range("AC" & case_row.Row)), Val(wb_bis.Worksheets("Case").Range("AD" & case_row.Row)), Val(wb_bis.Worksheets("Case").Range("AE" & case_row.Row)))
            action_date = DateSerial(Val(wb_bis.Worksheets("Case").Range("AF" & case_row.Row)), Val(wb_bis.Worksheets("Case").Range("AG" & case_row.Row)), Val(wb_bis.Worksheets("Case").Range("AH" & case_row.Row)))
            
            ' Determine what date to use for Most Recent Action - If Recert From date is greater than Application date
     '       If most_recent_action_date > appl_date Then
     '           If most_recent_action_date > time_zero Then _
     '             wb_sch.Worksheets("MA Workbook").Range("G22") = most_recent_action_date 'Most recent action date
     '       Else
     '           If appl_date > time_zero Then _
     '             wb_sch.Worksheets("MA Workbook").Range("G22") = appl_date 'Application date
     '       End If
            
            'put in dates for certification period
              wb_sch.Worksheets("MA Workbook").Range("F25") = open_date 'Most Recent Open Date
         
              wb_sch.Worksheets("MA Workbook").Range("F27") = action_date 'Most Recent Action Date


            
            'LIHEAP
      '      If wb_bis.Worksheets("Case").Range("AW" & case_row.Row) = "N" Then
      '          wb_sch.Worksheets("MA Workbook").Range("L1499") = "No"
      '      ElseIf wb_bis.Worksheets("Case").Range("AW" & case_row.Row) = "Y" Then
      '          wb_sch.Worksheets("MA Workbook").Range("L1499") = "Yes"
      '      End If
            
            'Date assigned - leave blank
            'wb_sch.Worksheets("MA Workbook").Range("G34") = wb_bis.Worksheets("Case").Cells(1, LastColumn_bis_case + 1)
            'wb_sch.Worksheets("MA Workbook").Range("G34") = Date
            
            'Fill in Absent Relatives in Section C of Workbook
    '            writerowabs = 25
    '            For j = 74 To 134 Step 30
    '                If Trim(wb_bis.Worksheets("Case").Cells(case_row.Row, j)) <> "" Then
    '                    writerowabs = writerowabs + 2
    '                    wb_sch.Worksheets("MA Workbook").Range("K" & writerowabs) = wb_bis.Worksheets("Case").Cells(case_row.Row, j + 1) & " " & wb_bis.Worksheets("Case").Cells(case_row.Row, j)
    '                    wb_sch.Worksheets("MA Workbook").Range("S" & writerowabs) = "LRR to " & wb_bis.Worksheets("Case").Cells(case_row.Row, j + 9)
    '                    wb_sch.Worksheets("MA Workbook").Range("V" & writerowabs) = wb_bis.Worksheets("Case").Cells(case_row.Row, j + 2)
    '                    wb_sch.Worksheets("MA Workbook").Range("Z" & writerowabs) = wb_bis.Worksheets("Case").Cells(case_row.Row, j + 3)
    '                    wb_sch.Worksheets("MA Workbook").Range("Z" & writerowabs + 1) = wb_bis.Worksheets("Case").Cells(case_row.Row, j + 5) & ", " & wb_bis.Worksheets("Case").Cells(case_row.Row, j + 6)
    '                Else
    '                    Exit For
    '                End If
    '            Next j
            
            'Start of Individual File processing
            birthdate = True
            school = ""
            schoolEM = ""
            schoolNM = ""
            citizen = ""
            noncitizenQ = ""
            noncitizenNOT = ""
            household = ""
            householdSNAP = ""
            write_row_citizen = 121
            write_row_earned = 7
            write_row_unearned = 17
            household_size = 0
            allcitizen = 1
            write_row = 10
            mandatory = ""
            Age = ""
            health = ""
            student = ""
            employment = ""
            Other = ""
            EligibleABAWD = ""
            ExemptABAWD = ""
            RSDI_LNS = ""
            VA_LNS = ""
            SSI_LNS = ""
            SSP_LNS = ""
            UC_LNS = ""
            WC_LNS = ""
            DI_LNS = ""
           ' wb_sch.Worksheets("MA Workbook").Shapes("CB 104").OLEFormat.Object.value = 1 'element 151
           ' wb_sch.Worksheets("MA Workbook").Shapes("CB 125").OLEFormat.Object.value = 1 'element 160
           ' wb_sch.Worksheets("MA Workbook").Shapes("CB 146").OLEFormat.Object.value = 1 'element 161
           ' wb_sch.Worksheets("MA Workbook").Shapes("CB 185").OLEFormat.Object.value = 1 'element 163
           ' wb_sch.Worksheets("MA Workbook").Shapes("CB 186").OLEFormat.Object.value = 1 'element 163
           ' wb_sch.Worksheets("MA Workbook").Shapes("CB 214").OLEFormat.Object.value = 1 'element 165
           ' wb_sch.Worksheets("MA Workbook").Shapes("CB 217").OLEFormat.Object.value = 1 'element 165
           ' wb_sch.Worksheets("MA Workbook").Shapes("CB 243").OLEFormat.Object.value = 1 'element 166
            
            'Loop through the Individual File
            For i = bis_ind_start_row To bis_ind_stop_row
            ' Fill in household info in section B
                write_row = write_row + 1
                If write_row < 23 Then 'only room for 12 names on workbook
                    wb_sch.Worksheets("MA Workbook").Range("J" & write_row) = lnf(wb_bis.Worksheets("Individual").Range("L" & i)) 'Line Number
                    wb_sch.Worksheets("MA Workbook").Range("L" & write_row) = Trim(wb_bis.Worksheets("Individual").Range("N" & i)) & " " & _
                        Trim(wb_bis.Worksheets("Individual").Range("P" & i)) & " " & Trim(wb_bis.Worksheets("Individual").Range("O" & i)) & " " & _
                        Trim(wb_bis.Worksheets("Individual").Range("Q" & i)) 'Full Name
                    wb_sch.Worksheets("MA Workbook").Range("AC" & write_row) = wb_bis.Worksheets("Individual").Range("J" & i) 'Individual Category
                    temp = wb_bis.Worksheets("Individual").Range("R" & i)
                    wb_sch.Worksheets("MA Workbook").Range("V" & write_row) = DateSerial(Val(Left(temp, 4)), Val(Mid(temp, 5, 2)), Val(Right(temp, 2)))
                    wb_sch.Worksheets("MA Workbook").Range("Y" & write_row) = wb_bis.Worksheets("Individual").Range("T" & i) 'Age
                    wb_sch.Worksheets("MA Workbook").Range("AA" & write_row) = wb_bis.Worksheets("Individual").Range("X" & i) 'Relationship
                    wb_sch.Worksheets("MA Workbook").Range("AE" & write_row) = wb_bis.Worksheets("Individual").Range("Z" & i) 'SSN
             '       If wb_bis.Worksheets("Individual").Range("Y" & i) = "NM" Then 'SNAP Not Received
             '           wb_sch.Worksheets("MA Workbook").Range("AJ" & write_row) = "No"
             '       Else
             '           wb_sch.Worksheets("MA Workbook").Range("AJ" & write_row) = "Yes" 'SNAP Received
             '           household_size = household_size + 1
             '       End If
                End If ' write_row < 23
            Next i

           
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'MA Schedule - Section 4
                i = bis_ind_start_row - 1
                For j = 51 To 73 Step 2
                    i = i + 1
                    wb_sch.Worksheets(review_number).Range("B" & j) = lnf(wb_bis.Worksheets("Individual").Range("L" & i)) 'Line number
          '          If Left(wb_bis.Worksheets("Individual").Range("Y" & i), 1) = "E" Then 'check SNAP Participation
          '              wb_sch.Worksheets(review_number).Range("E" & j) = "01"
          '          End If
                    Select Case UCase(wb_bis.Worksheets("Individual").Range("T" & i)) 'Relationship to Head of Household
                        Case "X  " 'Head of Household
                          If Val(wb_bis.Worksheets("Individual").Range("T" & i)) <= 19 Then
                            wb_sch.Worksheets(review_number).Range("N" & j) = "02"
                          Else
                            wb_sch.Worksheets(review_number).Range("N" & j) = "01"
                         End If
                        Case "W  ", "H  ", "CLH", "CLW" 'Spouse
                          If wb_bis.Worksheets("Individual").Range("T" & i) <= 19 Then
                            wb_sch.Worksheets(review_number).Range("N" & j) = "04"
                          Else
                            wb_sch.Worksheets(review_number).Range("N" & j) = "03"
                          End If
                        Case "F  ", "M  ", "SF ", "SM " 'Parent
                            wb_sch.Worksheets(review_number).Range("N" & j) = "05"
                        Case "D  ", "S  " 'Daughter, Son,
                            wb_sch.Worksheets(review_number).Range("N" & j) = "06"
                        Case "SS ", "SD " 'Stepson and stepdaughter
                            wb_sch.Worksheets(review_number).Range("N" & j) = "07"
                        Case "NR " 'Unrelated Person
                            wb_sch.Worksheets(review_number).Range("N" & j) = "20"
                        Case "GD ", "GS ", "GGS", "GGD"  'Grandchild or great grandchild
                            wb_sch.Worksheets(review_number).Range("N" & j) = "10"
                        Case Else 'Other related persons
                            wb_sch.Worksheets(review_number).Range("N" & j) = "14"
                    End Select
                    wb_sch.Worksheets(review_number).Range("R" & j) = wb_bis.Worksheets("Individual").Range("T" & i) 'Age
                    
                    'Gender
                    If wb_bis.Worksheets("Individual").Range("U" & i) = "F" Then
                        wb_sch.Worksheets(review_number).Range("V" & j) = "02"
                    ElseIf wb_bis.Worksheets("Individual").Range("U" & i) = "M" Then
                        wb_sch.Worksheets(review_number).Range("V" & j) = "01"
                    Else
                        wb_sch.Worksheets(review_number).Range("V" & j) = ""
                    End If
                    
                    'Race
                    Select Case wb_bis.Worksheets("Individual").Range("V" & i)
                        Case 1
                            wb_sch.Worksheets(review_number).Range("Y" & j) = 2
                        Case 3
                            wb_sch.Worksheets(review_number).Range("Y" & j) = 5
                        Case 4
                            wb_sch.Worksheets(review_number).Range("Y" & j) = 4
                        Case 5
                            wb_sch.Worksheets(review_number).Range("Y" & j) = 1
                        Case 6
                            wb_sch.Worksheets(review_number).Range("Y" & j) = 9
                        Case 7
                            wb_sch.Worksheets(review_number).Range("Y" & j) = 4
                        Case 8
                            If wb_bis.Worksheets("Individual").Range("W" & i) = 2 Then
                              wb_sch.Worksheets(review_number).Range("R" & j) = 3
                             Else
                              wb_sch.Worksheets(review_number).Range("R" & j) = 9
                            End If
                    End Select
                    
                    Select Case UCase(wb_bis.Worksheets("Individual").Range("AO" & i)) 'Citizenship
                        Case 1 'US Born
                            wb_sch.Worksheets(review_number).Range("AB" & j) = "01"
                        'Case 2 'Permanent Alien
                            'wb_sch.Worksheets(review_number).Range("AB" & j) = 2
                        'Case 3 'Temporary Alien
                            'wb_sch.Worksheets(review_number).Range("AB" & j) = 3
                        Case 4 'Refugee
                            wb_sch.Worksheets(review_number).Range("AB" & j) = "05"
                        'Case 5 'Illegal Alien
                            'wb_sch.Worksheets(review_number).Range("AB" & j) = 6
                        'Case 6 'Refugee Unaccompanied Minor
                            'wb_sch.Worksheets(review_number).Range("AB" & j) = 5
                    End Select
                    
                    'Education Level
           '         wb_sch.Worksheets(review_number).Range("V" & j) = wb_bis.Worksheets("Individual").Range("V" & i)
           '         'Change CIS education code from 98 to 0
           '         If Val(wb_sch.Worksheets(review_number).Range("V" & j)) = 98 Then
           '             wb_sch.Worksheets(review_number).Range("V" & j) = 0
           '         End If
           '         'Change CIS education code from 16 to 14
           '         If Val(wb_sch.Worksheets(review_number).Range("V" & j)) = 16 Then
           '             wb_sch.Worksheets(review_number).Range("V" & j) = "14"
           '         End If
                    
                 '  Select Case UCase(wb_bis.Worksheets("Individual").Range("AF" & i)) 'Employment Status
                        'Case 0  '
                        '    wb_sch.Worksheets(review_number).Range("Y" & j) = "01"
                 '       Case 1  'No Work History
                 '           wb_sch.Worksheets(review_number).Range("Y" & j) = 1
                        'Case 2 'Unemployed no work history within last 12 months
                            'wb_sch.Worksheets(review_number).Range("Y" & j) = 1
                        'Case 3 'Unemployed work history within last 12 months
                            'wb_sch.Worksheets(review_number).Range("Y" & j) = 1
                        'Case 4 'Unemployed on the job training
                            'wb_sch.Worksheets(review_number).Range("Y" & j) =
                        'Case 5 'Full-time employment
                            'wb_sch.Worksheets(review_number).Range("Y" & j) = 8
                        'Case 6 'Part-time employment 100 hours per month or more
                            'wb_sch.Worksheets(review_number).Range("Y" & j) = 8
                 '       Case 7  'Part-time employment less than 100 hours per month (unemployed)
                  '          wb_sch.Worksheets(review_number).Range("Y" & j) = 8
                        'Case 8 'Refused to comply with employment regulations
                            'wb_sch.Worksheets(review_number).Range("Y" & j) =
                        'Case 9 'Same as 6, but each 2 prev. mos. < 100 hrs next mo < 100 hrs(unemploy)
                            'wb_sch.Worksheets(review_number).Range("Y" & j) =
                 '       Case 10 'Self-employed
                 '          wb_sch.Worksheets(review_number).Range("Y" & j) = 7
                        'Case 11 'Participates in work study program
                            'wb_sch.Worksheets(review_number).Range("Y" & j) =
                        'Case 12 'Participates in VISTA Volunteer Program
                            'wb_sch.Worksheets(review_number).Range("Y" & j) =
                        'Case 13 'Part-time Self-employment
                            'wb_sch.Worksheets(review_number).Range("Y" & j) = 7
                 '   End Select
                    'wb_sch.Worksheets(review_number).Range("AA" & j) = wb_bis.Worksheets("Individual").Range(" " & i) 'Emp. Hours
                    'wb_sch.Worksheets(review_number).Range("AH" & j) = wb_bis.Worksheets("Individual").Range("Y" & i) 'ABAWD
                    If i = bis_ind_stop_row Then
                        Exit For
                    End If
            Next j 'end of MA Schedule - section 4
    
    End Select

End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'SNAP Computation
            'Rent
 '               If wb_bis.Worksheets("Case").Range("AU" & case_row.Row) <> 0 Then
            '       wb_sch.Worksheets("FS Computation").Range("B42") = wb_bis.Worksheets("Case").Range("AU" & case_row.Row)
  '                  wb_sch.Worksheets("MA Workbook").Range("N1462") = wb_bis.Worksheets("Case").Range("AU" & case_row.Row)
   '             End If
            'Standard Utility Determination
   '             Select Case wb_bis.Worksheets("Case").Range("AX" & case_row.Row)
   '                 Case "N", "U"
                    '    wb_sch.Worksheets("FS Computation").Range("B44") = "Non Heating "
   '                     wb_sch.Worksheets("MA Workbook").Range("O1489") = "Non Heating"
   '                 Case "H"
                    '    wb_sch.Worksheets("FS Computation").Range("B44") = "Heating "
   '                     wb_sch.Worksheets("MA Workbook").Range("O1489") = "Heating"
   '                 Case "L"
                    '    wb_sch.Worksheets("FS Computation").Range("B44") = "Limited "
   '                     wb_sch.Worksheets("MA Workbook").Range("O1489") = "Limited"
   '                 Case "P"
                    '    wb_sch.Worksheets("FS Computation").Range("B44") = "Telephone "
   '                     wb_sch.Worksheets("MA Workbook").Range("O1489") = "Telephone"
   '             End Select
                        
   ' End Select
    
'End Sub

Function lnf(linenum As Integer)
    lnf = WorksheetFunction.Text(linenum, "00")
End Function



