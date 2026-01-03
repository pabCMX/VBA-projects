Attribute VB_Name = "Populate_snap_pos_delimited_mod"
Sub Populatenew()

Dim wb_sch As Workbook, review_number As String, case_row As Range

income_freq = Array(0, 1, 4, 2, 2, 1, 0.5, 0.333333, 0.166667, 0.083333)

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
       ' If Val(wb_bis.Worksheets("Individual").Range("C" & i)) = review_number Then
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
    If bis_ind_stop_row > bis_ind_start_row Then
        'wb_bis.Worksheets("Individual").Range(Cells(bis_ind_start_row, 1), Cells(bis_ind_stop_row, LastColumn_bis_ind)).Sort key1:=wb_bis.Worksheets("Individual").Range("T" & bis_ind_start_row & ":T" & bis_ind_stop_row), order1:=xlDescending, Header:=xlNo
        'wb_bis.Worksheets("Individual").Range(Cells(bis_ind_start_row + 1, 1), Cells(bis_ind_stop_row, LastColumn_bis_ind)).Sort key1:=Range("L" & bis_ind_start_row + 1 & ":L" & bis_ind_stop_row), order1:=xlDescending, Header:=xlNo

        wb_bis.Worksheets("Individual").Sort.SortFields.Clear
        wb_bis.Worksheets("Individual").Sort.SortFields.Add Key:=Range( _
            "T" & bis_ind_start_row & ":T" & bis_ind_stop_row), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
            xlSortNormal
        With wb_bis.Worksheets("Individual").Sort
        .SetRange Range(Cells(bis_ind_start_row, 1), Cells(bis_ind_stop_row, LastColumn_bis_ind))
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
        End With
    
        wb_bis.Worksheets("Individual").Sort.SortFields.Clear
        wb_bis.Worksheets("Individual").Sort.SortFields.Add Key:=Range( _
            "L" & bis_ind_start_row + 1 & ":L" & bis_ind_stop_row), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
            xlSortNormal
        With wb_bis.Worksheets("Individual").Sort
        .SetRange Range(Cells(bis_ind_start_row + 1, 1), Cells(bis_ind_stop_row, LastColumn_bis_ind))
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
        End With
    End If
    
    'Put in worksheet data from case record
    Select Case program
    
        Case "FS Positive", "FS Supplemental"
            
            wb_sch.Worksheets("FS Workbook").Range("D16") = _
                wb_bis.Worksheets("Case").Range("AJ" & case_row.Row) 'Telephone Number
            
            time_zero = DateSerial(0, 0, 1) 'Time format with all zeros except day = 1
            
            temp = wb_bis.Worksheets("Case").Range("M" & case_row.Row) 'application date
            
            'if application date is zero, make application date = time zero
            If temp = 0 Then
                appl_date = time_zero
            Else
                appl_date = DateSerial(Val(Left(temp, 4)), Val(Mid(temp, 5, 2)), Val(Right(temp, 2)))
            End If
            
            If appl_date > time_zero Then _
              wb_sch.Worksheets("FS Workbook").Range("G20") = appl_date 'Application Date
            
            'get dates from delimited file
            most_recent_action_date = DateSerial(2000 + Val(wb_bis.Worksheets("Case").Range("X" & case_row.Row)), Val(wb_bis.Worksheets("Case").Range("Y" & case_row.Row)), Val(wb_bis.Worksheets("Case").Range("Z" & case_row.Row)))
            recert_from_date = DateSerial(2000 + Val(wb_bis.Worksheets("Case").Range("X" & case_row.Row)), Val(wb_bis.Worksheets("Case").Range("Y" & case_row.Row)), 1)
            recert_thru_date = DateSerial(2000 + Val(wb_bis.Worksheets("Case").Range("AA" & case_row.Row)), Val(wb_bis.Worksheets("Case").Range("AB" & case_row.Row)), Val(wb_bis.Worksheets("Case").Range("AC" & case_row.Row)))
            
            ' Determine what date to use for Most Recent Action - If Recert From date is greater than Application date
            If most_recent_action_date > appl_date Then
                If most_recent_action_date > time_zero Then _
                  wb_sch.Worksheets("FS Workbook").Range("G22") = most_recent_action_date 'Most recent action date
            Else
                If appl_date > time_zero Then _
                  wb_sch.Worksheets("FS Workbook").Range("G22") = appl_date 'Application date
            End If
            
            'put in dates for certification period
            If recert_from_date > time_zero Then _
              wb_sch.Worksheets("FS Workbook").Range("H24") = recert_from_date 'Certification From
            If recert_thru_date > time_zero Then _
              wb_sch.Worksheets("FS Workbook").Range("H25") = recert_thru_date 'Certification Thru
            wb_sch.Worksheets("FS Workbook").Range("R42") = wb_bis.Worksheets("Case").Range("AK" & case_row.Row) 'Coupon Amount
            'wb_sch.Worksheets("FS Workbook").Range("J936") = recert_from_date & " - " & recert_thru_date
            
            'LIHEAP
            If wb_bis.Worksheets("Case").Range("AW" & case_row.Row) = "N" Then
                wb_sch.Worksheets("FS Workbook").Range("L1499") = "No"
            ElseIf wb_bis.Worksheets("Case").Range("AW" & case_row.Row) = "Y" Then
                wb_sch.Worksheets("FS Workbook").Range("L1499") = "Yes"
            End If
            
            'Date assigned - leave blank
            'wb_sch.Worksheets("FS Workbook").Range("G34") = wb_bis.Worksheets("Case").Cells(1, LastColumn_bis_case + 1)
            'wb_sch.Worksheets("FS Workbook").Range("G34") = Date
            
            'Fill in Absent Relatives in Section C of Workbook
                writerowabs = 25
                For j = 74 To 134 Step 30
                    If Trim(wb_bis.Worksheets("Case").Cells(case_row.Row, j)) <> "" Then
                        writerowabs = writerowabs + 2
                        wb_sch.Worksheets("FS Workbook").Range("K" & writerowabs) = wb_bis.Worksheets("Case").Cells(case_row.Row, j + 1) & " " & wb_bis.Worksheets("Case").Cells(case_row.Row, j)
                        wb_sch.Worksheets("FS Workbook").Range("S" & writerowabs) = "LRR to " & wb_bis.Worksheets("Case").Cells(case_row.Row, j + 9)
                        wb_sch.Worksheets("FS Workbook").Range("V" & writerowabs) = wb_bis.Worksheets("Case").Cells(case_row.Row, j + 2)
                        wb_sch.Worksheets("FS Workbook").Range("Z" & writerowabs) = wb_bis.Worksheets("Case").Cells(case_row.Row, j + 3)
                        wb_sch.Worksheets("FS Workbook").Range("Z" & writerowabs + 1) = wb_bis.Worksheets("Case").Cells(case_row.Row, j + 5) & ", " & wb_bis.Worksheets("Case").Cells(case_row.Row, j + 6)
                    Else
                        Exit For
                    End If
                Next j
            
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
            wb_sch.Worksheets("FS Workbook").Shapes("CB 104").OLEFormat.Object.Value = 1 'element 151
            wb_sch.Worksheets("FS Workbook").Shapes("CB 125").OLEFormat.Object.Value = 1 'element 160
            wb_sch.Worksheets("FS Workbook").Shapes("CB 146").OLEFormat.Object.Value = 1 'element 161
            wb_sch.Worksheets("FS Workbook").Shapes("CB 185").OLEFormat.Object.Value = 1 'element 163
            wb_sch.Worksheets("FS Workbook").Shapes("CB 186").OLEFormat.Object.Value = 1 'element 163
            wb_sch.Worksheets("FS Workbook").Shapes("CB 214").OLEFormat.Object.Value = 1 'element 165
            wb_sch.Worksheets("FS Workbook").Shapes("CB 217").OLEFormat.Object.Value = 1 'element 165
            wb_sch.Worksheets("FS Workbook").Shapes("CB 243").OLEFormat.Object.Value = 1 'element 166
            
            'Loop through the Individual File
            For i = bis_ind_start_row To bis_ind_stop_row
            ' Fill in household info in section B
                write_row = write_row + 1
                If write_row < 23 Then 'only room for 12 names on workbook
                    wb_sch.Worksheets("FS Workbook").Range("K" & write_row) = lnf(wb_bis.Worksheets("Individual").Range("L" & i)) 'Line Number
                    wb_sch.Worksheets("FS Workbook").Range("M" & write_row) = Trim(wb_bis.Worksheets("Individual").Range("N" & i)) & " " & _
                        Trim(wb_bis.Worksheets("Individual").Range("P" & i)) & " " & Trim(wb_bis.Worksheets("Individual").Range("O" & i)) & " " & _
                        Trim(wb_bis.Worksheets("Individual").Range("Q" & i)) 'Full Name
                    wb_sch.Worksheets("FS Workbook").Range("W" & write_row) = wb_bis.Worksheets("Individual").Range("J" & i) 'Individual Category
                    temp = wb_bis.Worksheets("Individual").Range("R" & i)
                    wb_sch.Worksheets("FS Workbook").Range("Y" & write_row) = DateSerial(Val(Left(temp, 4)), Val(Mid(temp, 5, 2)), Val(Right(temp, 2)))
                    wb_sch.Worksheets("FS Workbook").Range("AB" & write_row) = wb_bis.Worksheets("Individual").Range("S" & i) 'Age
                    wb_sch.Worksheets("FS Workbook").Range("AD" & write_row) = wb_bis.Worksheets("Individual").Range("T" & i) 'Relationship
                    wb_sch.Worksheets("FS Workbook").Range("AF" & write_row) = wb_bis.Worksheets("Individual").Range("U" & i) 'SSN
                    If wb_bis.Worksheets("Individual").Range("Y" & i) = "NM" Then 'SNAP Not Received
                        wb_sch.Worksheets("FS Workbook").Range("AJ" & write_row) = "No"
                    Else
                        wb_sch.Worksheets("FS Workbook").Range("AJ" & write_row) = "Yes" 'SNAP Received
                        household_size = household_size + 1
                    End If
                End If ' write_row < 23
                
                'Element 110 - looking to see if every line number has a birthdate
                If wb_bis.Worksheets("Individual").Range("R" & i) = "" Or Val(wb_bis.Worksheets("Individual").Range("R" & i)) = 0 Then
                    birthdate = False
                End If
                'Element 111 - looking to see if older than 18 and in school
                If wb_bis.Worksheets("Individual").Range("S" & i) >= 18 And wb_bis.Worksheets("Individual").Range("W" & i) <> 0 Then
                    If Len(school) > 0 Then
                         school = school & ", " & lnf(wb_bis.Worksheets("Individual").Range("L" & i))
                    Else
                        school = school & lnf(wb_bis.Worksheets("Individual").Range("L" & i))
                    End If
                End If
                'Element 111 - looking to see if older than 18 and in school and eligiblity status = EM
                If wb_bis.Worksheets("Individual").Range("S" & i) >= 18 And wb_bis.Worksheets("Individual").Range("W" & i) <> 0 _
                And wb_bis.Worksheets("Individual").Range("Y" & i) = "EM" Then
                    If Len(schoolEM) > 0 Then
                         schoolEM = schoolEM & ", " & lnf(wb_bis.Worksheets("Individual").Range("L" & i))
                    Else
                        schoolEM = schoolEM & lnf(wb_bis.Worksheets("Individual").Range("L" & i))
                    End If
                End If
                 'Element 111 - looking to see if older than 18 and in school and eligiblity status = NM
                If wb_bis.Worksheets("Individual").Range("S" & i) >= 18 And wb_bis.Worksheets("Individual").Range("W" & i) <> 0 _
                And wb_bis.Worksheets("Individual").Range("Y" & i) = "NM" Then
                    If Len(schoolEM) > 0 Then
                         schoolNM = schoolNM & ", " & lnf(wb_bis.Worksheets("Individual").Range("L" & i))
                    Else
                        schoolNM = schoolNM & lnf(wb_bis.Worksheets("Individual").Range("L" & i))
                    End If
                End If
                'Element 130 for citizenship
                If wb_bis.Worksheets("Individual").Range("AF" & i) = 1 Then
                    If Len(citizen) > 0 Then
                        citizen = citizen & ", " & lnf(wb_bis.Worksheets("Individual").Range("L" & i))
                    Else
                        citizen = citizen & lnf(wb_bis.Worksheets("Individual").Range("L" & i))
                    End If
                End If
                'Element 130 for Qualified noncitizen
                If wb_bis.Worksheets("Individual").Range("AF" & i) <> 1 And Left(wb_bis.Worksheets("Individual").Range("Y" & i), 1) = "E" Then
                allcitizen = 0
                    If Len(noncitizenQ) > 0 Then
                        noncitizenQ = noncitizenQ & ", " & lnf(wb_bis.Worksheets("Individual").Range("L" & i))
                    Else
                        noncitizenQ = noncitizenQ & lnf(wb_bis.Worksheets("Individual").Range("L" & i))
                    End If
                End If
                'Element 130 for Not qualified noncitizen
                If wb_bis.Worksheets("Individual").Range("AF" & i) <> 1 And Left(wb_bis.Worksheets("Individual").Range("Y" & i), 1) = "N" Then
                allcitizen = 0
                    If Len(noncitizenNOT) > 0 Then
                        noncitizenNOT = noncitizenNOT & ", " & lnf(wb_bis.Worksheets("Individual").Range("L" & i))
                    Else
                        noncitizenNOT = noncitizenNOT & lnf(wb_bis.Worksheets("Individual").Range("L" & i))
                    End If
                End If
                'Element 150 - looking to see if eligiblity status = EM, EB, EW, or NM
                If wb_bis.Worksheets("Individual").Range("Y" & i) = "EM" Or wb_bis.Worksheets("Individual").Range("Y" & i) = "EB" _
                Or wb_bis.Worksheets("Individual").Range("Y" & i) = "EW" Or wb_bis.Worksheets("Individual").Range("Y" & i) = "NM" Then
                    If Len(household) > 0 Then
                        household = household & ", " & lnf(wb_bis.Worksheets("Individual").Range("L" & i))
                    Else
                        household = household & lnf(wb_bis.Worksheets("Individual").Range("L" & i))
                    End If
                End If
                'Element 150 - looking to see if eligiblity status = EM, EB, EW
                If wb_bis.Worksheets("Individual").Range("Y" & i) = "EM" Or wb_bis.Worksheets("Individual").Range("Y" & i) = "EB" _
                Or wb_bis.Worksheets("Individual").Range("Y" & i) = "EW" Then
                    If Len(householdSNAP) > 0 Then
                        householdSNAP = householdSNAP & ", " & lnf(wb_bis.Worksheets("Individual").Range("L" & i))
                    Else
                        householdSNAP = householdSNAP & lnf(wb_bis.Worksheets("Individual").Range("L" & i))
                    End If
                End If
                 'Element 150 - looking to see if household has elderly or disable person
                If wb_bis.Worksheets("Individual").Range("AE" & i) = 2 Or wb_bis.Worksheets("Individual").Range("AE" & i) = 3 Then
                    wb_sch.Worksheets("FS Workbook").Shapes("CB 285").OLEFormat.Object.Value = 1
                End If
                'Element 151 - looking to see if no one disqualified per CR/CIS
                If wb_bis.Worksheets("Individual").Range("Y" & i) = "DS" Or wb_bis.Worksheets("Individual").Range("Y" & i) = "DF" Then
                    wb_sch.Worksheets("FS Workbook").Shapes("CB 104").OLEFormat.Object.Value = 0
                End If
                'Element 160 - looking to see if all household members are exempt
                If wb_bis.Worksheets("Individual").Range("AE" & i) <> 30 Then
                    wb_sch.Worksheets("FS Workbook").Shapes("CB 125").OLEFormat.Object.Value = 0
                End If
                
                Select Case wb_bis.Worksheets("Individual").Range("AE" & i)
                
                'Element 160 - ETP code chart (mandatory)
                    Case 30, 40
                        If Len(mandatory) > 0 Then
                          mandatory = mandatory & ", " & lnf(wb_bis.Worksheets("Individual").Range("L" & i))
                        Else
                           mandatory = mandatory & lnf(wb_bis.Worksheets("Individual").Range("L" & i))
                        End If
                'Element 160 - ETP code chart (age)
                    Case 1, 2
                        If Len(Age) > 0 Then
                            Age = Age & ", " & lnf(wb_bis.Worksheets("Individual").Range("L" & i))
                        Else
                            Age = Age & lnf(wb_bis.Worksheets("Individual").Range("L" & i))
                        End If
                'Element 160 - ETP code chart (health)
                    Case 3
                        If Len(health) > 0 Then
                            health = health & ", " & lnf(wb_bis.Worksheets("Individual").Range("L" & i))
                        Else
                            health = health & lnf(wb_bis.Worksheets("Individual").Range("L" & i))
                        End If
                'Element 160 - ETP code chart (employment)
                    Case 17
                        If Len(employment) > 0 Then
                           employment = employment & ", " & lnf(wb_bis.Worksheets("Individual").Range("L" & i))
                        Else
                            employment = employment & lnf(wb_bis.Worksheets("Individual").Range("L" & i))
                        End If
                'Element 160 - ETP code chart (student)
                    Case 20
                        If Len(student) > 0 Then
                           student = student & ", " & lnf(wb_bis.Worksheets("Individual").Range("L" & i))
                        Else
                            student = student & lnf(wb_bis.Worksheets("Individual").Range("L" & i))
                        End If
                'Element 160 - ETP code chart (student with ETP code 13)
                    Case 13
                        If (wb_bis.Worksheets("Individual").Range("AB" & i) = 16 Or wb_bis.Worksheets("Individual").Range("AB" & i) = 17) _
                        And wb_bis.Worksheets("Individual").Range("w" & i) <> 0 Then
                            If Len(student) > 0 Then
                                student = student & ", " & lnf(wb_bis.Worksheets("Individual").Range("L" & i))
                            Else
                                student = student & lnf(wb_bis.Worksheets("Individual").Range("L" & i))
                            End If
                        End If
                'Element 160 - ETP code chart (other)
                    Case Else
                        If Len(Other) > 0 Then
                           Other = Other & ", " & lnf(wb_bis.Worksheets("Individual").Range("L" & i))
                        Else
                            Other = Other & lnf(wb_bis.Worksheets("Individual").Range("L" & i))
                        End If
                    End Select
                'Element 161 - looking to see if ABAWD in household
                If wb_bis.Worksheets("Individual").Range("Y" & i) = "EB" Then
                    wb_sch.Worksheets("FS Workbook").Shapes("CB 146").OLEFormat.Object.Value = 0
                    wb_sch.Worksheets("FS Workbook").Shapes("CB 185").OLEFormat.Object.Value = 0
                    wb_sch.Worksheets("FS Workbook").Shapes("CB 214").OLEFormat.Object.Value = 0
                    wb_sch.Worksheets("FS Workbook").Shapes("CB 243").OLEFormat.Object.Value = 0
                Else
                    wb_sch.Worksheets("FS Workbook").Shapes("CB 317").OLEFormat.Object.Value = 1
                    Select Case wb_bis.Worksheets("Individual").Range("AE" & i)
                        Case 4, 6, 10, 14, 15, 16, 17, 18, 19, 20, 21
                            If Len(ExemptABAWD) > 0 Then
                                ExemptABAWD = ExemptABAWD & ", " & lnf(wb_bis.Worksheets("Individual").Range("L" & i))
                            Else
                                ExemptABAWD = ExemptABAWD & lnf(wb_bis.Worksheets("Individual").Range("L" & i))
                            End If
                        Case Else
                            If Len(EligibleABAWD) > 0 Then
                                EligibleABAWD = EligibleABAWD & ", " & lnf(wb_bis.Worksheets("Individual").Range("L" & i))
                            Else
                                EligibleABAWD = EligibleABAWD & lnf(wb_bis.Worksheets("Individual").Range("L" & i))
                            End If
                        End Select
                End If
                'Element 163
                If wb_bis.Worksheets("Individual").Range("AC" & i) = 21 Or wb_bis.Worksheets("Individual").Range("AC" & i) = 22 _
                Or wb_bis.Worksheets("Individual").Range("AC" & i) = 23 Then
                    wb_sch.Worksheets("FS Workbook").Shapes("CB 186").OLEFormat.Object.Value = 0
                End If
                'Element 165 - 2nd check box
                If wb_bis.Worksheets("Individual").Range("AN" & i) <> 0 Then
                    wb_sch.Worksheets("FS Workbook").Shapes("CB 217").OLEFormat.Object.Value = 0
                End If
                'Element 165 - 3rd check box
                If wb_bis.Worksheets("Individual").Range("AN" & i) = 1 Or wb_bis.Worksheets("Individual").Range("AN" & i) = 3 Then
                    wb_sch.Worksheets("FS Workbook").Shapes("CB 223").OLEFormat.Object.Value = 1
                End If
                'Element 165 - 4th check box
                 If wb_bis.Worksheets("Individual").Range("AN" & i) = 2 Or wb_bis.Worksheets("Individual").Range("AN" & i) = 4 Then
                    wb_sch.Worksheets("FS Workbook").Shapes("CB 240").OLEFormat.Object.Value = 1
                End If
                'Element 166 - 2nd check box
                Select Case wb_bis.Worksheets("Individual").Range("AO" & i)
                    Case 5, 6, 7, 10, 13
                        wb_sch.Worksheets("FS Workbook").Shapes("CB 225").OLEFormat.Object.Value = 1
                    End Select
                'Element 166 - 4th check box
                Select Case wb_bis.Worksheets("Individual").Range("AB" & i)
                    Case 21, 22, 23
                        wb_sch.Worksheets("FS Workbook").Shapes("CB 246").OLEFormat.Object.Value = 1
                    End Select
                'Element 170 - 5th check box
                 If wb_bis.Worksheets("Individual").Range("AC" & i) = 24 Then
                    wb_sch.Worksheets("FS Workbook").Shapes("CB 230").OLEFormat.Object.Value = 1
                End If
                'Element 331 - RSDI Benefits
                 If wb_bis.Worksheets("Individual").Range("BX" & i) > 0 Then
                    If Len(RSDI_LNS) > 0 Then
                        RSDI_LNS = RSDI_LNS & ", " & lnf(wb_bis.Worksheets("Individual").Range("L" & i))
                    Else
                        RSDI_LNS = RSDI_LNS & lnf(wb_bis.Worksheets("Individual").Range("L" & i))
                    End If
                End If
                'Element 332 - Veteran's Benefits
                 If wb_bis.Worksheets("Individual").Range("BZ" & i) > 0 Then
                    If Len(VA_LNS) > 0 Then
                        VA_LNS = VA_LNS & ", " & lnf(wb_bis.Worksheets("Individual").Range("L" & i))
                    Else
                        VA_LNS = VA_LNS & lnf(wb_bis.Worksheets("Individual").Range("L" & i))
                    End If
                End If
                'Element 333 - SSI Benefits
                 If wb_bis.Worksheets("Individual").Range("CB" & i) > 0 Then
                    If Len(SSI_LNS) > 0 Then
                        SSI_LNS = SSI_LNS & ", " & lnf(wb_bis.Worksheets("Individual").Range("L" & i))
                    Else
                        SSI_LNS = SSI_LNS & lnf(wb_bis.Worksheets("Individual").Range("L" & i))
                    End If
                End If
                'Element 333 - SSP Benefits
                 If wb_bis.Worksheets("Individual").Range("CI" & i) > 0 Then
                    If Len(SSP_LNS) > 0 Then
                        SSP_LNS = SSP_LNS & ", " & lnf(wb_bis.Worksheets("Individual").Range("L" & i))
                    Else
                        SSP_LNS = SSP_LNS & lnf(wb_bis.Worksheets("Individual").Range("L" & i))
                    End If
                End If
                 'Element 334 - UC Benefits
                 If wb_bis.Worksheets("Individual").Range("CD" & i) > 0 Then
                    If Len(UC_LNS) > 0 Then
                        UC_LNS = UC_LNS & ", " & lnf(wb_bis.Worksheets("Individual").Range("L" & i))
                    Else
                        UC_LNS = UC_LNS & lnf(wb_bis.Worksheets("Individual").Range("L" & i))
                    End If
                End If
                'Element 335 - WC Benefits
                 If wb_bis.Worksheets("Individual").Range("CF" & i) > 0 Then
                    If Len(WC_LNS) > 0 Then
                        WC_LNS = WC_LNS & ", " & lnf(wb_bis.Worksheets("Individual").Range("L" & i))
                    Else
                        WC_LNS = WC_LNS & lnf(wb_bis.Worksheets("Individual").Range("L" & i))
                    End If
                End If
                'Element 342 child Support
                If wb_bis.Worksheets("Individual").Range("BI" & i) > 0 Then
                    wb_sch.Worksheets("FS Workbook").Shapes("CB 746").OLEFormat.Object.Value = 1
                End If
                'Element 343 - Deemed Income
                 If wb_bis.Worksheets("Individual").Range("CH" & i) > 0 Then
                    If Len(DI_LNS) > 0 Then
                        DI_LNS = DI_LNS & ", " & lnf(wb_bis.Worksheets("Individual").Range("L" & i))
                    Else
                        DI_LNS = DI_LNS & lnf(wb_bis.Worksheets("Individual").Range("L" & i))
                    End If
                End If
                
                
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                'SNAP Computation sheet
                
                'wb_sch.Worksheets("FS Computation").Range("B11") = wb_bis.Worksheets("Case").Range("W" & i) 'Number of Individuals
                'wb_sch.Worksheets("FS Computation").Range("B11") = household_size
                'If write_row_earned < 10 Then
                'Earned Income - 11 - regular income
                'If wb_bis.Worksheets("Individual").Range("BR" & i) <> 0 And write_row_earned < 10 Then
                '   write_row_earned = write_row_earned + 1
                '   wb_sch.Worksheets("FS Computation").Range("A" & write_row_earned) = lnf(wb_bis.Worksheets("Individual").Range("L" & i)) & "/11"
                '   monthly_income = Round(income_freq(wb_bis.Worksheets("Individual").Range("BQ" & i)) * wb_bis.Worksheets("Individual").Range("BR" & i), 0)
                '   wb_sch.Worksheets("FS Computation").Range("B" & write_row_earned) = monthly_income
                'End If
                'Earned Income - 12 - self employment
                'If wb_bis.Worksheets("Individual").Range("BS" & i) <> 0 And write_row_earned < 10 Then
                '   write_row_earned = write_row_earned + 1
                '   wb_sch.Worksheets("FS Computation").Range("A" & write_row_earned) = lnf(wb_bis.Worksheets("Individual").Range("L" & i)) & "/12"
                '   monthly_income = Round(income_freq(wb_bis.Worksheets("Individual").Range("BQ" & i)) * wb_bis.Worksheets("Individual").Range("BS" & i), 0)
                '   wb_sch.Worksheets("FS Computation").Range("B" & write_row_earned) = monthly_income
                'End If
                'End If 'write_row_earned < 10
                
                'If write_row_unearned < 22 Then
                'Unearned Income - 31 - SSD/RSDI
                'If wb_bis.Worksheets("Individual").Range("BX" & i) <> 0 And write_row_unearned < 22 Then
                '   write_row_unearned = write_row_unearned + 1
                '   wb_sch.Worksheets("FS Computation").Range("A" & write_row_unearned) = lnf(wb_bis.Worksheets("Individual").Range("L" & i)) & "/31"
                '   monthly_income = Round(income_freq(wb_bis.Worksheets("Individual").Range("BW" & i)) * wb_bis.Worksheets("Individual").Range("BX" & i), 0)
                '   wb_sch.Worksheets("FS Computation").Range("B" & write_row_unearned) = monthly_income
                'End If
                'Unearned Income - 32 - VA
                'If wb_bis.Worksheets("Individual").Range("BZ" & i) <> 0 And write_row_unearned < 22 Then
                '   write_row_unearned = write_row_unearned + 1
                '   wb_sch.Worksheets("FS Computation").Range("A" & write_row_unearned) = lnf(wb_bis.Worksheets("Individual").Range("L" & i)) & "/32"
                '   monthly_income = Round(income_freq(wb_bis.Worksheets("Individual").Range("BY" & i)) * wb_bis.Worksheets("Individual").Range("BZ" & i), 0)
                '   wb_sch.Worksheets("FS Computation").Range("B" & write_row_unearned) = monthly_income
                'End If
                'Unearned Income - 33 - SSI
                'If wb_bis.Worksheets("Individual").Range("CB" & i) <> 0 And write_row_unearned < 22 Then
                '   write_row_unearned = write_row_unearned + 1
                '   wb_sch.Worksheets("FS Computation").Range("A" & write_row_unearned) = lnf(wb_bis.Worksheets("Individual").Range("L" & i)) & "/33"
                '   monthly_income = Round(income_freq(wb_bis.Worksheets("Individual").Range("CA" & i)) * wb_bis.Worksheets("Individual").Range("CB" & i), 0)
                '   wb_sch.Worksheets("FS Computation").Range("B" & write_row_unearned) = monthly_income
                'End If
                'Unearned Income - 34 - Unemployment
                'If wb_bis.Worksheets("Individual").Range("CD" & i) <> 0 And write_row_unearned < 22 Then
                '   write_row_unearned = write_row_unearned + 1
                '   wb_sch.Worksheets("FS Computation").Range("A" & write_row_unearned) = lnf(wb_bis.Worksheets("Individual").Range("L" & i)) & "/34"
                '   monthly_income = Round(income_freq(wb_bis.Worksheets("Individual").Range("CC" & i)) * wb_bis.Worksheets("Individual").Range("CD" & i), 0)
                '   wb_sch.Worksheets("FS Computation").Range("B" & write_row_unearned) = monthly_income
                'End If
                'Unearned Income - 35 - Workman Comp
                'If wb_bis.Worksheets("Individual").Range("CF" & i) <> 0 And write_row_unearned < 22 Then
                '   write_row_unearned = write_row_unearned + 1
                '   wb_sch.Worksheets("FS Computation").Range("A" & write_row_unearned) = lnf(wb_bis.Worksheets("Individual").Range("L" & i)) & "/35"
                '   monthly_income = Round(income_freq(wb_bis.Worksheets("Individual").Range("CE" & i)) * wb_bis.Worksheets("Individual").Range("CF" & i), 0)
                '   wb_sch.Worksheets("FS Computation").Range("B" & write_row_unearned) = monthly_income
                'End If
                'Unearned Income - 43 - Deemed Income
                'If wb_bis.Worksheets("Individual").Range("CH" & i) <> 0 And write_row_unearned < 22 Then
                '   write_row_unearned = write_row_unearned + 1
                '   wb_sch.Worksheets("FS Computation").Range("A" & write_row_unearned) = lnf(wb_bis.Worksheets("Individual").Range("L" & i)) & "/43"
                '   monthly_income = Round(income_freq(wb_bis.Worksheets("Individual").Range("CG" & i)) * wb_bis.Worksheets("Individual").Range("CH" & i), 0)
                '   wb_sch.Worksheets("FS Computation").Range("B" & write_row_unearned) = monthly_income
                'End If
                'Unearned Income - 44 - Public Assistance
                'If wb_bis.Worksheets("Individual").Range("CJ" & i) <> 0 And write_row_unearned < 22 Then
               '    write_row_unearned = write_row_unearned + 1
                '   wb_sch.Worksheets("FS Computation").Range("A" & write_row_unearned) = lnf(wb_bis.Worksheets("Individual").Range("L" & i)) & "/44"
              '     monthly_income = Round(income_freq(wb_bis.Worksheets("Individual").Range("CI" & i)) * wb_bis.Worksheets("Individual").Range("CJ" & i), 0)
               '    wb_sch.Worksheets("FS Computation").Range("B" & write_row_unearned) = monthly_income
                'End If
                 'Unearned Income - 45 - Educational grants
                'If wb_bis.Worksheets("Individual").Range("CL" & i) <> 0 And write_row_unearned < 22 Then
                 '  write_row_unearned = write_row_unearned + 1
                 '  wb_sch.Worksheets("FS Computation").Range("A" & write_row_unearned) = lnf(wb_bis.Worksheets("Individual").Range("L" & i)) & "/45"
                '   monthly_income = Round(income_freq(wb_bis.Worksheets("Individual").Range("CK" & i)) * wb_bis.Worksheets("Individual").Range("CL" & i), 0)
                '   wb_sch.Worksheets("FS Computation").Range("B" & write_row_unearned) = monthly_income
                'End If
                 'Unearned Income - 24, 26, 32 - Child Support - add all CS incomes under code 50
                'If (wb_bis.Worksheets("Individual").Range("CV" & i) <> 0 Or _
                '    wb_bis.Worksheets("Individual").Range("CX" & i) <> 0 Or _
                '    wb_bis.Worksheets("Individual").Range("CZ" & i) <> 0) And write_row_unearned < 22 Then
                '   write_row_unearned = write_row_unearned + 1
                '   wb_sch.Worksheets("FS Computation").Range("A" & write_row_unearned) = lnf(wb_bis.Worksheets("Individual").Range("L" & i)) & "/50"
                '   monthly_income = Round(income_freq(wb_bis.Worksheets("Individual").Range("CU" & i)) * wb_bis.Worksheets("Individual").Range("CV" & i), 0) + _
                '        Round(income_freq(wb_bis.Worksheets("Individual").Range("CW" & i)) * wb_bis.Worksheets("Individual").Range("CX" & i), 0) + _
                '        Round(income_freq(wb_bis.Worksheets("Individual").Range("CY" & i)) * wb_bis.Worksheets("Individual").Range("CZ" & i), 0)
                '   wb_sch.Worksheets("FS Computation").Range("B" & write_row_unearned) = monthly_income
                'End If
                'End If ' if write row unearned < 22
           Next i 'End of Main Individual File Loop
           
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'SNAP Schedule - Section 4
                i = bis_ind_start_row - 1
                For j = 89 To 122 Step 3
                    i = i + 1
                    wb_sch.Worksheets(review_number).Range("B" & j) = lnf(wb_bis.Worksheets("Individual").Range("L" & i)) 'Line number
                    If Left(wb_bis.Worksheets("Individual").Range("Y" & i), 1) = "E" Then 'check SNAP Participation
                        wb_sch.Worksheets(review_number).Range("E" & j) = "01"
                    End If
                    Select Case UCase(wb_bis.Worksheets("Individual").Range("T" & i)) 'Relationship to Head of Household
                        Case "X  " 'Head of Household
                            wb_sch.Worksheets(review_number).Range("H" & j) = 1
                        Case "W  ", "H  ", "CLH", "CLW" 'Spouse
                            wb_sch.Worksheets(review_number).Range("H" & j) = 2
                        Case "F  ", "M  ", "SF ", "SM " 'Parent
                            wb_sch.Worksheets(review_number).Range("H" & j) = 3
                        Case "D  ", "SD ", "S  ", "SS " 'Daughter, Stepdaughter, Son, Stepson
                            wb_sch.Worksheets(review_number).Range("H" & j) = 4
                        Case "NR " 'Unrelated Person
                            wb_sch.Worksheets(review_number).Range("H" & j) = 6
                        Case Else 'Other related persons
                            wb_sch.Worksheets(review_number).Range("H" & j) = 5
                    End Select
                    wb_sch.Worksheets(review_number).Range("J" & j) = wb_bis.Worksheets("Individual").Range("S" & i) 'Age
                    
                    'Gender
                    If wb_bis.Worksheets("Individual").Range("CO" & i) = "F" Then
                        wb_sch.Worksheets(review_number).Range("M" & j) = "02"
                    ElseIf wb_bis.Worksheets("Individual").Range("CO" & i) = "M" Then
                        wb_sch.Worksheets(review_number).Range("M" & j) = "01"
                    Else
                        wb_sch.Worksheets(review_number).Range("M" & j) = ""
                    End If
                    
                    'Race
                    Select Case wb_bis.Worksheets("Individual").Range("CP" & i)
                        Case 1
                            wb_sch.Worksheets(review_number).Range("P" & j) = "05"
                        Case 3
                            wb_sch.Worksheets(review_number).Range("P" & j) = "03"
                        Case 4
                            wb_sch.Worksheets(review_number).Range("P" & j) = "04"
                        Case 5
                            wb_sch.Worksheets(review_number).Range("P" & j) = "07"
                        Case 6
                            wb_sch.Worksheets(review_number).Range("P" & j) = "12"
                        Case 7
                            wb_sch.Worksheets(review_number).Range("P" & j) = "06"
                        Case 8
                            wb_sch.Worksheets(review_number).Range("P" & j) = "99"
                    End Select
                    
                    Select Case UCase(wb_bis.Worksheets("Individual").Range("AF" & i)) 'Citizenship
                        Case 1 'US Born
                            wb_sch.Worksheets(review_number).Range("S" & j) = "01"
                        'Case 2 'Permanent Alien
                            'wb_sch.Worksheets(review_number).Range("S" & j) = 2
                        'Case 3 'Temporary Alien
                            'wb_sch.Worksheets(review_number).Range("S" & j) = 3
                        Case 4 'Refugee
                            wb_sch.Worksheets(review_number).Range("S" & j) = "05"
                        'Case 5 'Illegal Alien
                            'wb_sch.Worksheets(review_number).Range("S" & j) = 6
                        'Case 6 'Refugee Unaccompanied Minor
                            'wb_sch.Worksheets(review_number).Range("S" & j) = 5
                    End Select
                    
                    'Education Level
                    wb_sch.Worksheets(review_number).Range("V" & j) = wb_bis.Worksheets("Individual").Range("V" & i)
                    'Change CIS education code from 98 to 0
                    If Val(wb_sch.Worksheets(review_number).Range("V" & j)) = 98 Then
                        wb_sch.Worksheets(review_number).Range("V" & j) = "00"
                    End If
                    'Change CIS education code from 16 to 14
                    If Val(wb_sch.Worksheets(review_number).Range("V" & j)) = 16 Then
                        wb_sch.Worksheets(review_number).Range("V" & j) = "14"
                    End If
                    
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
                Next j 'end of SNAP Schedule - section 4
                
            'Element 110 - Check if birthdates are in CIS
            If birthdate = True Then
                wb_sch.Worksheets("FS Workbook").Shapes("CB 25").OLEFormat.Object.Value = 1
            End If
            
            wb_sch.Worksheets("FS Workbook").Range("I110") = school 'Element 111 - if older than 18 and in school
            wb_sch.Worksheets("FS Workbook").Range("I112") = schoolEM 'Element 111 - if older than 18 & in school & status=EM
            wb_sch.Worksheets("FS Workbook").Range("I114") = schoolNM 'Element 111 - if older than 18 & in school & status=EM
            
            'Element 130
            If allcitizen = 1 Then
               write_row_citizen = write_row_citizen + 1
               wb_sch.Worksheets("FS Workbook").Range("G" & write_row_citizen) = "All"
               wb_sch.Worksheets("FS Workbook").Range("I" & write_row_citizen) = "Citizen"
            Else
            If Len(citizen) > 0 Then
                write_row_citizen = write_row_citizen + 1
                wb_sch.Worksheets("FS Workbook").Range("G" & write_row_citizen) = citizen
                wb_sch.Worksheets("FS Workbook").Range("I" & write_row_citizen) = "Citizen"
            End If
            If Len(noncitizenQ) > 0 Then
                write_row_citizen = write_row_citizen + 1
                wb_sch.Worksheets("FS Workbook").Range("G" & write_row_citizen) = noncitizenQ
                wb_sch.Worksheets("FS Workbook").Range("I" & write_row_citizen) = "Non-Citizen"
            End If
            If Len(noncitizenNOT) > 0 Then
                write_row_citizen = write_row_citizen + 1
                wb_sch.Worksheets("FS Workbook").Range("G" & write_row_citizen) = noncitizenNOT
                wb_sch.Worksheets("FS Workbook").Range("I" & write_row_citizen) = "Non-Citizen"
            End If
            End If
                
            'Element 140 - checking to see if person from PA
            If wb_bis.Worksheets("Case").Range("AH" & case_row.Row) = "PA" Then
                wb_sch.Worksheets("FS Workbook").Shapes("CB 56").OLEFormat.Object.Value = 1
            End If
          
            wb_sch.Worksheets("FS Workbook").Range("I213") = household 'Element 150 - eligiblity EB, EM, EW, or NM
            wb_sch.Worksheets("FS Workbook").Range("I215") = householdSNAP 'Element 150 - eligiblity EB, EM, or EW
            
            Dim lnarray As Variant, exarray As Variant

            lnarray = Array(267, 269, 271, 273, 275, 276)
            exarray = Array(268, 270, 272, 274, 275, 276)

            'Element 160
            write_row_ETP = -1
            If Len(mandatory) > 0 Then
                write_row_ETP = write_row_ETP + 1
                wb_sch.Worksheets("FS Workbook").Range("G" & lnarray(write_row_ETP)) = mandatory
                wb_sch.Worksheets("FS Workbook").Range("L" & exarray(write_row_ETP)) = "mandatory"
            End If
             If Len(Age) > 0 Then
                write_row_ETP = write_row_ETP + 1
                wb_sch.Worksheets("FS Workbook").Range("G" & lnarray(write_row_ETP)) = Age
                wb_sch.Worksheets("FS Workbook").Range("L" & exarray(write_row_ETP)) = "age"
            End If
             If Len(health) > 0 Then
                write_row_ETP = write_row_ETP + 1
                wb_sch.Worksheets("FS Workbook").Range("G" & lnarray(write_row_ETP)) = health
                wb_sch.Worksheets("FS Workbook").Range("L" & exarray(write_row_ETP)) = "health"
            End If
             If Len(student) > 0 Then
                write_row_ETP = write_row_ETP + 1
                wb_sch.Worksheets("FS Workbook").Range("G" & lnarray(write_row_ETP)) = student
                wb_sch.Worksheets("FS Workbook").Range("L" & exarray(write_row_ETP)) = "student"
            End If
             If Len(employment) > 0 Then
                write_row_ETP = write_row_ETP + 1
                wb_sch.Worksheets("FS Workbook").Range("G" & lnarray(write_row_ETP)) = employment
                wb_sch.Worksheets("FS Workbook").Range("L" & exarray(write_row_ETP)) = "employment"
            End If
             If Len(Other) > 0 Then
                write_row_ETP = write_row_ETP + 1
                wb_sch.Worksheets("FS Workbook").Range("G" & lnarray(write_row_ETP)) = Other
                wb_sch.Worksheets("FS Workbook").Range("L" & exarray(write_row_ETP)) = "other"
            End If
            
            'Element 161
            If wb_sch.Worksheets("FS Workbook").Shapes("CB 146").OLEFormat.Object.Value = 0 Then
                wb_sch.Worksheets("FS Workbook").Range("M344") = EligibleABAWD
                wb_sch.Worksheets("FS Workbook").Range("M348") = ExemptABAWD
            'Element 163
                If Len(EligibleABAWD) > 0 And Len(ExemptABAWD) > 0 Then
                    wb_sch.Worksheets("FS Workbook").Range("N370") = EligibleABAWD & ", " & ExemptABAWD
                ElseIf Len(EligibleABAWD) = 0 And Len(ExemptABAWD) = 0 Then
                    wb_sch.Worksheets("FS Workbook").Range("N370") = ""
                ElseIf Len(EligibleABAWD) > 0 And Len(ExemptABAWD) = 0 Then
                    wb_sch.Worksheets("FS Workbook").Range("N370") = EligibleABAWD
                ElseIf Len(EligibleABAWD) = 0 And Len(ExemptABAWD) > 0 Then
                    wb_sch.Worksheets("FS Workbook").Range("N370") = ExemptABAWD
                End If
            End If
            'Element 311
            'wb_sch.Worksheets("FS Workbook").Range("J934") = recert_from_date & " - " & recert_thru_date
            'RSDI_LNS
            'Element 331 - RSDI Line Numbers
            If Len(RSDI_LNS) > 0 Then
                wb_sch.Worksheets("FS Workbook").Range("K1087") = RSDI_LNS
            End If
            'Element 332 - VA Line Numbers
            If Len(VA_LNS) > 0 Then
                wb_sch.Worksheets("FS Workbook").Range("N1096") = VA_LNS
            End If
            'Element 333 - SSI Line Numbers
            If Len(SSI_LNS) > 0 Then
                wb_sch.Worksheets("FS Workbook").Range("J1109") = SSI_LNS
            End If
            'Element 333 - SSP Line Numbers
            If Len(SSP_LNS) > 0 Then
                wb_sch.Worksheets("FS Workbook").Range("J1113") = SSP_LNS
            End If
            'Element 334 - UC Line Numbers
            If Len(UC_LNS) > 0 Then
                wb_sch.Worksheets("FS Workbook").Range("N1124") = UC_LNS
            End If
            'Element 335 - WC Line Numbers
            If Len(WC_LNS) > 0 Then
                wb_sch.Worksheets("FS Workbook").Range("N1198") = WC_LNS
            End If
            'Element 343 - Deemed Income Line Numbers
            If Len(DI_LNS) > 0 Then
                wb_sch.Worksheets("FS Workbook").Range("N1235") = WC_LNS
            End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'SNAP Computation
            'Rent
                If wb_bis.Worksheets("Case").Range("AU" & case_row.Row) <> 0 Then
            '       wb_sch.Worksheets("FS Computation").Range("B42") = wb_bis.Worksheets("Case").Range("AU" & case_row.Row)
                    wb_sch.Worksheets("FS Workbook").Range("N1462") = wb_bis.Worksheets("Case").Range("AU" & case_row.Row)
                End If
            'Standard Utility Determination
                Select Case wb_bis.Worksheets("Case").Range("AX" & case_row.Row)
                    Case "N", "U"
                    '    wb_sch.Worksheets("FS Computation").Range("B44") = "Non Heating "
                        wb_sch.Worksheets("FS Workbook").Range("O1489") = "Non Heating"
                    Case "H"
                    '    wb_sch.Worksheets("FS Computation").Range("B44") = "Heating "
                        wb_sch.Worksheets("FS Workbook").Range("O1489") = "Heating"
                    Case "L"
                    '    wb_sch.Worksheets("FS Computation").Range("B44") = "Limited "
                        wb_sch.Worksheets("FS Workbook").Range("O1489") = "Limited"
                    Case "P"
                    '    wb_sch.Worksheets("FS Computation").Range("B44") = "Telephone "
                        wb_sch.Worksheets("FS Workbook").Range("O1489") = "Telephone"
                End Select
                        
    End Select
    
End Sub

Function lnf(linenum As Integer)
    lnf = WorksheetFunction.Text(linenum, "00")
End Function

