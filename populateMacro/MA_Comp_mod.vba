Attribute VB_Name = "MA_Comp_mod"
Public columnletter As String
Sub MA_Comp_finalresults()

    Range("C8") = Range(columnletter & "8")
    Range("C144") = Range(columnletter & "144")
    
    'copy category and status code
    Select Case columnletter
    Case "D", "E", "F"
        Range("C6") = Range("E6")
        Range("C7") = Range("E7")
    Case "H", "I", "J"
        Range("C6") = Range("I6")
        Range("C7") = Range("I7")
    Case "K", "L", "M"
        Range("C6") = Range("L6")
        Range("C7") = Range("L7")
    End Select
    
    For i = 10 To 15
        Range("C" & i) = Range(columnletter & i)
    Next i

    For i = 17 To 21
        Range("C" & i) = Range(columnletter & i)
    Next i
    
    For i = 23 To 27
        Range("C" & i) = Range(columnletter & i)
    Next i

    For i = 29 To 33
        Range("C" & i) = Range(columnletter & i)
    Next i

    For i = 35 To 39
        Range("C" & i) = Range(columnletter & i)
    Next i

    For i = 41 To 45
        Range("C" & i) = Range(columnletter & i)
    Next i

    For i = 47 To 51
        Range("C" & i) = Range(columnletter & i)
    Next i

    For i = 55 To 56
        Range("C" & i) = Range(columnletter & i)
    Next i

    For i = 58 To 59
        Range("C" & i) = Range(columnletter & i)
    Next i

    For i = 61 To 62
        Range("C" & i) = Range(columnletter & i)
    Next i

    For i = 64 To 65
        Range("C" & i) = Range(columnletter & i)
    Next i
    
    For i = 71 To 72
        Range("C" & i) = Range(columnletter & i)
    Next i

    For i = 76 To 80
        Range("C" & i) = Range(columnletter & i)
    Next i
    
    For i = 84 To 85
        Range("C" & i) = Range(columnletter & i)
    Next i

    For i = 87 To 88
        Range("C" & i) = Range(columnletter & i)
    Next i
    
    For i = 90 To 91
        Range("C" & i) = Range(columnletter & i)
    Next i
    
    For i = 97 To 102
        Range("C" & i) = Range(columnletter & i)
    Next i

    For i = 104 To 110
        Range("C" & i) = Range(columnletter & i)
    Next i
    
    For i = 114 To 119
        Range("C" & i) = Range(columnletter & i)
    Next i
    
    For i = 121 To 127
        Range("C" & i) = Range(columnletter & i)
    Next i
    
    For i = 132 To 133
        Range("C" & i) = Range(columnletter & i)
    Next i
    
End Sub
    
Sub show_formMAC2()
    UserFormMAC2.Show
End Sub

Sub redisplayformMAC2()
    UserFormMAC2.Show
End Sub
Sub show_formMAC3()
    UserFormMAC3.Show
End Sub

Sub redisplayformMAC3()
    UserFormMAC3.Show
End Sub
Sub MA_Comp_finalresults3()

    Range("G8") = Range(columnletter & "8")
    Range("G144") = Range(columnletter & "144")
    
    'copy category and status code
    Select Case columnletter
    Case "D", "E", "F"
        Range("G6") = Range("E6")
        Range("G7") = Range("E7")
    Case "H", "I", "J"
        Range("G6") = Range("I6")
        Range("G7") = Range("I7")
    Case "K", "L", "M"
        Range("G6") = Range("L6")
        Range("G7") = Range("L7")
    End Select
    
    For i = 10 To 15
        Range("G" & i) = Range(columnletter & i)
    Next i

    For i = 17 To 21
        Range("G" & i) = Range(columnletter & i)
    Next i
    
    For i = 23 To 27
        Range("G" & i) = Range(columnletter & i)
    Next i

    For i = 29 To 33
        Range("G" & i) = Range(columnletter & i)
    Next i

    For i = 35 To 39
        Range("G" & i) = Range(columnletter & i)
    Next i

    For i = 41 To 45
        Range("G" & i) = Range(columnletter & i)
    Next i

    For i = 47 To 51
        Range("G" & i) = Range(columnletter & i)
    Next i

    For i = 55 To 56
        Range("G" & i) = Range(columnletter & i)
    Next i

    For i = 58 To 59
        Range("G" & i) = Range(columnletter & i)
    Next i

    For i = 61 To 62
        Range("G" & i) = Range(columnletter & i)
    Next i

    For i = 64 To 65
        Range("G" & i) = Range(columnletter & i)
    Next i
    
    For i = 71 To 72
        Range("G" & i) = Range(columnletter & i)
    Next i

    For i = 76 To 80
        Range("G" & i) = Range(columnletter & i)
    Next i
    
    For i = 84 To 85
        Range("G" & i) = Range(columnletter & i)
    Next i

    For i = 87 To 88
        Range("G" & i) = Range(columnletter & i)
    Next i
    
    For i = 90 To 91
        Range("G" & i) = Range(columnletter & i)
    Next i
    
    For i = 97 To 102
        Range("G" & i) = Range(columnletter & i)
    Next i

    For i = 104 To 110
        Range("G" & i) = Range(columnletter & i)
    Next i
    
    For i = 114 To 119
        Range("G" & i) = Range(columnletter & i)
    Next i
    
    For i = 121 To 127
        Range("G" & i) = Range(columnletter & i)
    Next i
    
    For i = 132 To 133
        Range("G" & i) = Range(columnletter & i)
    Next i
    
End Sub
Sub MAClearButtons1to15()

  For i = 1 To 15
    ActiveSheet.Shapes("OB " & i).Select
    With Selection
        .Value = xlOff
    End With
  Next i
  Range("AI102").Select
End Sub
Sub MAClearButtons16to24()

  For i = 16 To 24
    ActiveSheet.Shapes("OB " & i).Select
    With Selection
        .Value = xlOff
    End With
  Next i
    Range("AI115").Select
End Sub

Sub MAClearButtons25to27()

  For i = 25 To 27
    ActiveSheet.Shapes("OB " & i).Select
    With Selection
        .Value = xlOff
    End With
  Next i
    
    Range("AI305").Select
End Sub

Sub MAClearButtons28to33()

  For i = 28 To 33
    ActiveSheet.Shapes("OB " & i).Select
    With Selection
        .Value = xlOff
    End With
  Next i
    Range("AI422").Select
End Sub

Sub MAClearButtons34to45()

  For i = 34 To 45
    ActiveSheet.Shapes("OB " & i).Select
    With Selection
        .Value = xlOff
    End With
  Next i
    Range("AI521").Select
End Sub

Sub MAClearButtons46to57()

  For i = 46 To 57
    ActiveSheet.Shapes("OB " & i).Select
    With Selection
        .Value = xlOff
    End With
  Next i
    Range("AI643").Select
End Sub


Sub MAClearButtons58to66()

  For i = 58 To 66
    ActiveSheet.Shapes("OB " & i).Select
    With Selection
        .Value = xlOff
    End With
  Next i
    Range("AI754").Select
End Sub


Sub MAClearButtons67to78()

  For i = 67 To 78
    ActiveSheet.Shapes("OB " & i).Select
    With Selection
        .Value = xlOff
    End With
  Next i
    Range("AI851").Select
End Sub

Sub MAClearButtons79to93()

  For i = 79 To 93
    ActiveSheet.Shapes("OB " & i).Select
    With Selection
        .Value = xlOff
    End With
  Next i
    Range("AI967").Select
End Sub
Sub MAClearButtons94to105()

  For i = 94 To 105
    ActiveSheet.Shapes("OB " & i).Select
    With Selection
        .Value = xlOff
    End With
  Next i
    Range("AI1076").Select
End Sub

Sub MAClearButtons106to117()

  For i = 106 To 117
    ActiveSheet.Shapes("OB " & i).Select
    With Selection
        .Value = xlOff
    End With
  Next i
    Range("AI1192").Select
End Sub

Sub MAClearButtons118to126()

  For i = 118 To 126
    ActiveSheet.Shapes("OB " & i).Select
    With Selection
        .Value = xlOff
    End With
  Next i
    Range("AI1305").Select
End Sub

Sub MAClearButtons127to132()

  For i = 127 To 132
    ActiveSheet.Shapes("OB " & i).Select
    With Selection
        .Value = xlOff
    End With
  Next i
    Range("AI1408").Select
End Sub

Sub resources_ex()

If ActiveSheet.Shapes("CB 354").OLEFormat.Object.Value = 1 Then
    ActiveSheet.Shapes("CB 353").OLEFormat.Object.Value = 1
    ActiveSheet.Shapes("CB 700").OLEFormat.Object.Value = 1
    ActiveSheet.Shapes("CB 703").OLEFormat.Object.Value = 1
    ActiveSheet.Shapes("CB 708").OLEFormat.Object.Value = 1
    ActiveSheet.Shapes("CB 707").OLEFormat.Object.Value = 1
    ActiveSheet.Shapes("CB 729").OLEFormat.Object.Value = 1
    ActiveSheet.Shapes("CB 730").OLEFormat.Object.Value = 1
    ActiveSheet.Shapes("CB 537").OLEFormat.Object.Value = 1
    ActiveSheet.Shapes("CB 536").OLEFormat.Object.Value = 1
    ActiveSheet.Shapes("CB 3249").OLEFormat.Object.Value = 1
    ActiveSheet.Shapes("CB 3245").OLEFormat.Object.Value = 1
    ActiveSheet.Shapes("CB 3268").OLEFormat.Object.Value = 1
    ActiveSheet.Shapes("CB 273").OLEFormat.Object.Value = 1
    ActiveSheet.Shapes("CB 4519").OLEFormat.Object.Value = 1
    ActiveSheet.Shapes("CB 4520").OLEFormat.Object.Value = 1
   ' ActiveSheet.Shapes("OB 34").OLEFormat.Object.Value = 1
   ' ActiveSheet.Shapes("OB 37").OLEFormat.Object.Value = 1
   ' ActiveSheet.Shapes("OB 40").OLEFormat.Object.Value = 1
   ' ActiveSheet.Shapes("OB 43").OLEFormat.Object.Value = 1
   ' ActiveSheet.Shapes("OB 46").OLEFormat.Object.Value = 1
   ' ActiveSheet.Shapes("OB 49").OLEFormat.Object.Value = 1
   ' ActiveSheet.Shapes("OB 55").OLEFormat.Object.Value = 1
   ' ActiveSheet.Shapes("OB 52").OLEFormat.Object.Value = 1
    Range("G922").Select
Else
   ActiveSheet.Shapes("CB 353").OLEFormat.Object.Value = 0
    ActiveSheet.Shapes("CB 700").OLEFormat.Object.Value = 0
    ActiveSheet.Shapes("CB 703").OLEFormat.Object.Value = 0
    ActiveSheet.Shapes("CB 708").OLEFormat.Object.Value = 0
    ActiveSheet.Shapes("CB 707").OLEFormat.Object.Value = 0
    ActiveSheet.Shapes("CB 729").OLEFormat.Object.Value = 0
    ActiveSheet.Shapes("CB 730").OLEFormat.Object.Value = 0
    ActiveSheet.Shapes("CB 537").OLEFormat.Object.Value = 0
    ActiveSheet.Shapes("CB 536").OLEFormat.Object.Value = 0
    ActiveSheet.Shapes("CB 3249").OLEFormat.Object.Value = 0
    ActiveSheet.Shapes("CB 3245").OLEFormat.Object.Value = 0
    ActiveSheet.Shapes("CB 3268").OLEFormat.Object.Value = 0
    ActiveSheet.Shapes("CB 273").OLEFormat.Object.Value = 0
    ActiveSheet.Shapes("CB 4519").OLEFormat.Object.Value = 0
    ActiveSheet.Shapes("CB 4520").OLEFormat.Object.Value = 0
  '  ActiveSheet.Shapes("OB 34").OLEFormat.Object.Value = 0
  '  ActiveSheet.Shapes("OB 37").OLEFormat.Object.Value = 0
  '  ActiveSheet.Shapes("OB 40").OLEFormat.Object.Value = 0
  '  ActiveSheet.Shapes("OB 43").OLEFormat.Object.Value = 0
  '  ActiveSheet.Shapes("OB 46").OLEFormat.Object.Value = 0
  '  ActiveSheet.Shapes("OB 49").OLEFormat.Object.Value = 0
  '  ActiveSheet.Shapes("OB 55").OLEFormat.Object.Value = 0
  '  ActiveSheet.Shapes("OB 52").OLEFormat.Object.Value = 0
End If
End Sub
Sub ScheduleClean2()
'
'
'
Dim i As Long, j As Long, k As Long, l As Long, m As Long, n As Long, o As Long
Dim Member As Long, Person As Long
Dim Relationship As String, Gender As String, Category As String
Dim Age As String
'ask if user wants to proceed
msg = "Data will be updated using the entries on the Workbook and Comp sheet." & vbCrLf & vbCrLf & _
"Please check all entries after process is finished." & vbCrLf & vbCrLf & "Do you want to proceed?"
Ans = MsgBox(msg, vbYesNo + vbExclamation)
If Ans = vbNo Then End
'
'Blocks that are completed without conditions (besides the Empty condition)are completed with most often used codes.
'
'Section I
'
Application.ScreenUpdating = False

'
'6. Most likely used code (for completed reviews that are not dropped)
[F16] = "1"
'63. Code for a clean review
If [S16] = Empty Then
    [S16] = "01"
End If
'62a. Taken from coverage code on facesheet
If [I16] = Empty Then
    [I16] = Sheets("MA Workbook").[AC37]
End If
'62b. Taken from coverage code on facesheet
If [N16] = Empty Then
    [N16] = Sheets("MA Workbook").[AC37]
End If
'62ab. Taken from the category on comp sheet because everywhere else in the workbook has the grant group attached to the category.
[H17] = Sheets("MA Income Comp").[C12]
'62bb. Same as above
If [M17] = Empty Then
    [M17] = Sheets("MA Income Comp").[C12]
End If

'Section II
'9. Taken from facesheet
[A27] = Sheets("MA Workbook").[F25]
'10. Taken from facesheet
[L27] = Sheets("MA Workbook").[F25]
'11. Based on type of action entered on the facesheet.
If LCase(Sheets("MA Workbook").[F29]) = "open" Then
     [V27] = "1"
End If
'11. Based on type of action entered on the facesheet.
If LCase(Sheets("MA Workbook").[F29]) = "renewal" Then
  [V27] = "3"
End If
'9a. Based on 11.  If it was a renewal then it will be coded that they received prior assistance. Otherwise, it is left blank.
If [V27] = "3" Then
    [J27] = "1"
End If
'12. Counts the # of eligible line numbers from facesheet.
[Y27] = Application.WorksheetFunction.CountIf(Sheets("MA Workbook").[AK11:AK22], "<>")
'13. Taken from Resource Comp sheet.
[AB27] = Sheets("MA Resources").[B145]
'14-16. Usually zero.
[AH27, AN27, B32, X32] = "0"
'39. Taken from Workbook Element 371 (Combined Gross) Column 2.
[AF32] = Sheets("MA Workbook").[M1425] 'old [AF32] = Sheets("MA Workbook").[K1389]
'40. Taken from Workbook Element 372 (Combined Net) Column 2.
[AN32] = Sheets("MA Workbook").[M1430] '[AN32] = Sheets("MA Workbook").[K1391]

[AN32, AB27, AF32].NumberFormat = "0" 'To get rid of cents in the above cells.
'39-40. Puts zeros in if "No Income" is checked in Element 371.
If Sheets("MA Workbook").Shapes("Check Box 565").OLEFormat.Object.Value = 1 Then
    [AF32] = "0"
    [AN32] = "0"
End If


'Section III
'41. Fills in Line Numbers from facesheet. Changes format to a 2-digit number.
[B51].Select
For i = 11 To 20
Person = Sheets("MA Workbook").Cells(i, 10).Value
    If Person = Empty Then
        'ActiveCell.Value = Empty
        Exit For
    Else
        ActiveCell.Value = Person
    End If
ActiveCell.Offset(2, 0).Select
Next i
[B51:B73].Select
Selection.NumberFormat = "00"
'42a. Fills in FS code based on prompt.
msg = "Household receiving FS in the review month?"
Ans = MsgBox(msg, vbYesNo)
[F51].Select
For p = 11 To 20
    Person = Sheets("MA Workbook").Cells(p, 10).Value
    If Person = Empty Then
        'ActiveCell.Value = Empty
        Exit For
    Else
        If Ans = vbYes Then
            ActiveCell.Value = "2"
        ElseIf Ans = vbNo Then
            ActiveCell.Value = "3"
        End If
    End If
    ActiveCell.Offset(2, 0).Select
Next p
'42b. Fills in medical code based on info on facesheet:
[G51].Select
For s = 11 To 20
Rec = Sheets("MA Workbook").Cells(s, 36).Value 'Values from the "REC." column on the facesheet.
Categ = Sheets("MA Workbook").Cells(s, 29).Value 'Values from the "CAT." column on the facesheet.
    If Rec = "Y" Then
        If Categ = "C" Or Categ = "U" Then
            ActiveCell.Value = "1"
        ElseIf Categ = "J" Then
            ActiveCell.Value = "6"
        Else
            ActiveCell.Value = "3"
        End If
    ElseIf Rec = "N" Then
        ActiveCell.Value = "5"
    End If
   ActiveCell.Offset(2, 0).Select
Next s
'43. Based on coverage code & "ELIG." column, both on the facesheet.
[J51].Select
CC = Sheets("MA Workbook").[AC37]
For p = 11 To 20
Elig = Sheets("MA Workbook").Cells(p, 37).Value
    If Elig = Empty And Sheets("MA Workbook").Cells(p, 12).Value <> Empty Then
        ActiveCell.Value = "17"
    ElseIf Elig <> Empty Then
    Select Case CC
        Case "36", "37"
            ActiveCell.Value = "06"
        Case "12"
            ActiveCell.Value = "09"
        Case "35", "40", "43", "45", "49", "50", "51", "52", "63", "64"
            ActiveCell.Value = "11"
        Case "14", "23"
            ActiveCell.Value = "12"
        Case "81", "83", "26"
            ActiveCell.Value = "13"
        Case "69"
            ActiveCell.Value = "14"
        Case "27"
            ActiveCell.Value = "15"
        Case "82"
            ActiveCell.Value = "16"
    End Select
    End If
ActiveCell.Offset(2, 0).Select
Next p
'44. Based on values from the "REL. or SIG." column on the facesheet.
[N51].Select
For j = 11 To 20
Relationship = Sheets("MA Workbook").Cells(j, 27).Value
Select Case Relationship
            Case "X": ActiveCell.Value = "01"
            Case "S": ActiveCell.Value = "06"
            Case "D": ActiveCell.Value = "06"
            Case "M": ActiveCell.Value = "05"
            Case "F": ActiveCell.Value = "05"
            Case "H": ActiveCell.Value = "03"
            Case "W": ActiveCell.Value = "03"
            Case "SS": ActiveCell.Value = "07"
            Case "SD": ActiveCell.Value = "07"
            Case "GS": ActiveCell.Value = "10"
            Case "GD": ActiveCell.Value = "10"
            Case "B": ActiveCell.Value = "11"
            Case "SR": ActiveCell.Value = "11"
            Case "N": ActiveCell.Value = "11"
            Case "NC": ActiveCell.Value = "11"
            Case "GF": ActiveCell.Value = "11"
            Case "GM": ActiveCell.Value = "11"
            Case "NR"
                If Sheets("MA Workbook").Cells(j, 25).Value > 17 Then
                    ActiveCell.Value = "14"
                ElseIf Sheets("MA Workbook").Cells(j, 25).Value < 18 Then
                    ActiveCell.Value = "13"
                End If
End Select
ActiveCell.Offset(2, 0).Select
Next j
'45. Based on values from the "AGE" column on the facesheet.
[R51].Select
For k = 11 To 20
    Age = Sheets("MA Workbook").Cells(k, 25).Value
    If Age = Empty Then
        ActiveCell.Value = Empty
    ElseIf Age = "NB" Or Age = "0" Then
        ActiveCell.Value = "0"
    Else
        ActiveCell.Value = Age
    End If
ActiveCell.Offset(2, 0).Select
Next k
'46. Gender based on values from the "REL. or SIG." column on the facesheet. The X member must be left blank because there's no way to tell the gender from "X"
[V53].Select
For l = 12 To 20
    Gender = Sheets("MA Workbook").Cells(l, 27).Value
    Select Case Gender
            Case "X": ActiveCell.Value = ""
            Case "S": ActiveCell.Value = "1"
            Case "D": ActiveCell.Value = "2"
            Case "M": ActiveCell.Value = "2"
            Case "F": ActiveCell.Value = "1"
            Case "H": ActiveCell.Value = "1"
            Case "W": ActiveCell.Value = "2"
            Case "SS": ActiveCell.Value = "1"
            Case "SD": ActiveCell.Value = "2"
            Case "GS": ActiveCell.Value = "1"
            Case "GD": ActiveCell.Value = "2"
            Case "B": ActiveCell.Value = "1"
            Case "SR": ActiveCell.Value = "2"
            Case "N": ActiveCell.Value = "1"
            Case "NC": ActiveCell.Value = "2"
            Case "GF": ActiveCell.Value = "1"
            Case "GM": ActiveCell.Value = "2"
    End Select
    ActiveCell.Offset(2, 0).Select
Next l
'47. Race based on value entered after a propmt.  Enters same race for the entire household.
Race = InputBox("Race?" & vbNewLine & "1-White" & vbNewLine & "2-Black" & vbNewLine & "3-Hispanic" & vbNewLine & "4-Asian or Pacific Islander" & vbNewLine & "5-American Indian" & vbNewLine & "9-Unknown", "Race")
[Y51].Select
For q = 11 To 20
Person = Sheets("MA Workbook").Cells(q, 10).Value
    If Person = Empty Then
        ActiveCell.Value = Empty
    Else
        ActiveCell.Value = Race
    End If
ActiveCell.Offset(2, 0).Select
Next q
'49. Enters code for 'unknown'.
[AF51].Select
For r = 51 To 71 Step 2
Age49 = Cells(r, 18).Value
    If Age49 = Empty Then
        ActiveCell.Value = Empty
    Else
        ActiveCell.Value = "9"
    End If
ActiveCell.Offset(2, 0).Select
Next r
'50. Enters a hyphen.
[AI51].Select
For m = 11 To 20
Age50 = Sheets("MA Workbook").Cells(m, 25).Value
                If Age50 = Empty Then
                   ActiveCell.Value = Empty
                Else
                    ActiveCell.Value = "-"
                End If
ActiveCell.Offset(2, 0).Select
Next m
'51. Based on age from facesheet. This first part only fills for children and elderly.
[AM51].Select
For n = 11 To 20
    Age51 = Val(Sheets("MA Workbook").Cells(n, 25).Value)
    'if there are no more line numbers, exit for
    If Sheets("MA Workbook").Cells(n, 10).Value = Empty Then
        'ActiveCell.Value = Empty
        Exit For
    Else
        If Age51 >= 0 And Age51 < 16 Then
            ActiveCell.Value = "-"
        ElseIf Age51 > 64 Then
            ActiveCell.Value = "22"
        End If
    End If
    ActiveCell.Offset(2, 0).Select
Next n
'This part fill in a non recipient code based on values from 42b.
[AM51].Select
For t = 51 To 71 Step 2
Rec2 = Cells(t, 7).Value
    If ActiveCell = Empty And Rec2 = "6" Then
            ActiveCell.Value = "22"
    End If
ActiveCell.Offset(2, 0).Select
Next t
'52. Based on catgeory shown on schedule (this worksheet) and only completes for Line 01.
[AQ51].Select
For o = 51 To 67 Step 2
    Member = Cells(o, 2).Value
    Category = [Q10]
                If Member = "01" Then
                    Select Case Category
                        Case "PAN", "PJN", "TAN", "TJN"
                            ActiveCell.Value = "3"
                        Case "PAW", "PJW"
                            ActiveCell.Value = "8"
                        Case Else
                            ActiveCell.Value = "1"
                    End Select
                ElseIf Member <> Empty Then
                        ActiveCell.Value = "1"
                ElseIf Member = Empty Then
                        Exit For
                End If
    ActiveCell.Offset(2, 0).Select
Next o

' Section IV
' loops thru the comp sheet and gathers income from each line number
iread = 15
iwrite = 76
icol1 = -4
icol2 = 0
lastln = ""
Do Until Sheets("MA Income Comp").Range("A" & iread) = Empty Or iread = 20
    Ln = Left(Sheets("MA Income Comp").Range("A" & iread), 2)
    If Ln <> lastln Then
        iwrite = iwrite + 2
        Range("B" & iwrite) = Ln
        icol1 = -4
        icol2 = 0
    End If
    icol1 = icol1 + 10
    icol2 = icol2 + 10
    Cells(iwrite, icol1) = Right(Sheets("MA Income Comp").Range("A" & iread), 2)
    Cells(iwrite, icol2) = Sheets("MA Income Comp").Range("C" & iread)
    lastln = Ln
    iread = iread + 1
Loop
'Taken from info on the Income Comp sheet
'Dim Ln, Ln2, Ln3, Ln4 As String 'These are the "LN#/TYPE" lines from the Income Comp sheet, which must be entered in the Line#/Element format. ex: 01/311.
'Ln = Sheets("MA Income Comp").Range("A15")
'Ln2 = Sheets("MA Income Comp").Range("A16")
'Ln3 = Sheets("MA Income Comp").Range("A17")
'Ln4 = Sheets("MA Income Comp").Range("A18")
''53-57. Enters the line # and income source. Also takes into account multiple sources for the same line #.
'If Sheets("MA Income Comp").[A15] <> Empty Then
'    [B78] = Left(Ln, 2)
'    [F78] = Right(Ln, 2)
'    [J78] = Sheets("MA Income Comp").[C15]
'End If
'If Sheets("MA Income Comp").[A16] <> Empty And Left(Ln2, 2) = Left(Ln, 2) Then
'    [P78] = Right(Ln2, 2)
'    [T78] = Sheets("MA Income Comp").[C16]
'ElseIf Sheets("MA Income Comp").[A16] <> Empty Then
'    [B80] = Left(Ln2, 2)
'    [F80] = Right(Ln2, 2)
'    [J80] = Sheets("MA Income Comp").[C16]
'End If
'If Sheets("MA Income Comp").[A17] <> Empty And Left(Ln3, 2) = Left(Ln2, 2) Then
'    [P80] = Right(Ln3, 2)
'    [T80] = Sheets("MA Income Comp").[C17]
'ElseIf Sheets("MA Income Comp").[A17] <> Empty Then
'    [B82] = Left(Ln3, 2)
'    [F82] = Right(Ln3, 2)
'    [J82] = Sheets("MA Income Comp").[C17]
'End If
'If Sheets("MA Income Comp").[A18] <> Empty And Left(Ln4, 2) = Left(Ln3, 2) Then
'    [P82] = Right(Ln4, 2)
'    [T82] = Sheets("MA Income Comp").[C18]
'ElseIf Sheets("MA Income Comp").[A18] <> Empty Then
'    [B84] = Left(Ln4, 2)
'    [F84] = Right(Ln4, 2)
'    [J84] = Sheets("MA Income Comp").[C18]
'End If

' Changes format to a 2-digit number
[B78:B84].Select
Selection.NumberFormat = "00"

' 53-57. Whenever the Income Comp sheet is not completed, the income info is taken from Workbook Element 331 (RSDI) and/or Element 346 (Other) because these are the most likely the sources of income when the income sheet is not completed.
Ln5 = Sheets("MA Workbook").Range("N1089") 'Ln5 = Sheets("MA Workbook").Range("N1065")
Ln6 = Sheets("MA Workbook").Range("N1094") 'Ln6 = Sheets("MA Workbook").Range("N1070")
If Sheets("MA Income Comp").[A15] = Empty And Sheets("MA Workbook").[I1089] <> Empty Then
    [B78] = Right(Ln5, 2)
    [F78] = "31"
    [J78] = Sheets("MA Workbook").[I1089]
    If Sheets("MA Workbook").[I1094] <> Empty Then
        [B80] = Right(Ln6, 2)
        [F80] = "31"
        [J80] = Sheets("MA Workbook").[I1094]
    End If
End If
If Sheets("MA Income Comp").[A15] = Empty And Sheets("MA Workbook").[G1324] <> Empty Then
    If [B78] = Empty Then
        [B78] = Sheets("MA Workbook").[G1324]
        [F78] = "46"
        [J78] = Application.WorksheetFunction.Sum(Sheets("MA Workbook").[L1324], Sheets("MA Workbook").[L1326], Sheets("MA Workbook").[L1328])
    ElseIf [B78] <> Empty And [B80] = Empty Then
                [B80] = Sheets("MA Workbook").[G1324]
                [F80] = "46"
                [J80] = Application.WorksheetFunction.Sum(Sheets("MA Workbook").[L1324], Sheets("MA Workbook").[L1326], Sheets("MA Workbook").[L1328])
    End If
End If
            
'Section VI
'Citizenship.
Range("O118").Value = "1"
'Renewal Type based on value from 11.
    If Range("V27").Value = "1" Then
        Range("AE118").Value = "4"
    ElseIf Range("V27").Value = "3" Then
        Range("AE118").Value = "1"
    End If
'MSP/SSA Referral
Range("AJ118").Value = "0"
Application.ScreenUpdating = True


Range("BU4").Select
End Sub

Sub TPL_Fill()
'TPL
'Takes information from MA Workbook and MA Income Comp
'
Dim linea, wlinea, wlineaname, lineb, wlineb, wlinebname As String

[G11] = Sheets("MA Workbook").[F39]
[A13] = Sheets("MA Workbook").[D20]

'Medicare Claim # taken from Workbook
If Sheets("MA Workbook").[I1092] <> Empty Then
    [F19] = Sheets("MA Workbook").[I1092]
    [C19] = [A7]
    ActiveSheet.Shapes("Check Box 33").OLEFormat.Object.Value = 1
    ActiveSheet.Shapes("Check Box 31").OLEFormat.Object.Value = 1
End If

'#2 Category
ActiveSheet.Shapes("Check Box 3").OLEFormat.Object.Value = 1
'Question 1 in Column A
ActiveSheet.Shapes("Check Box 43").OLEFormat.Object.Value = 1
'Question 2 in Column A
ActiveSheet.Shapes("Check Box 13").OLEFormat.Object.Value = 1
ActiveSheet.Shapes("Check Box 11").OLEFormat.Object.Value = 1
ActiveSheet.Shapes("Check Box 12").OLEFormat.Object.Value = 1
'Employer, taken from wages section in "MA Income Comp". Only enters Employer 1 of first two line numbers
If Sheets("MA Income Comp").[A72] <> "Employer 1" Then
linea = Sheets("MA Income Comp").[A71]
wlinea = Right(linea, 2)
For i = 11 To 20
    If Sheets("MA Workbook").Cells(i, 10).Value = wlinea Then
    wlineaname = Sheets("MA Workbook").Cells(i, 12).Value
    End If
Next i
[C26] = wlineaname
[C29] = Sheets("MA Income Comp").[A72]
End If

If Sheets("MA Income Comp").[E72] <> "Employer 1" Then
lineb = Sheets("MA Income Comp").[E71]
wlineb = Right(lineb, 2)
For k = 11 To 20
    If Sheets("MA Workbook").Cells(k, 10).Value = wlineb Then
    wlinebname = Sheets("MA Workbook").Cells(k, 12).Value
    End If
Next k
[G26] = wlinebname
[G29] = Sheets("MA Income Comp").[E72]
End If

'Question 3 in Column A
ActiveSheet.Shapes("Check Box 23").OLEFormat.Object.Value = 1
'Question 4 in Column A
ActiveSheet.Shapes("Check Box 15").OLEFormat.Object.Value = 1
ActiveSheet.Shapes("Check Box 17").OLEFormat.Object.Value = 1
'Question 5 in Column A (Absent parent)taken from LRR section on the facesheet of the Workbook
If Sheets("MA Workbook").[L27] <> Empty Then
    ActiveSheet.Shapes("Check Box 18").OLEFormat.Object.Value = 1
    [C45] = "5. Name: " & Sheets("MA Workbook").[L27]
    [C46] = "Address: " & Sheets("MA Workbook").[AE27] & ", " & Sheets("MA Workbook").[AE28]
    [H45] = "SSN: " & Sheets("MA Workbook").[AA27]
End If

'Question 6 in Column A
ActiveSheet.Shapes("Check Box 21").OLEFormat.Object.Value = 1

Range("A19").Select

End Sub

Sub TPL_Clear()
'
'Clears everything that the TPL macro did (but nothing more).
'
'
msg = "Any information previously entered into the Third Party sheet by the above button will be cleared.  Are you sure you want to proceed?"
Ans = MsgBox(msg, vbYesNo)
If Ans = vbNo Then
    Exit Sub
End If

[G11:I11].ClearContents
[A13:A14].ClearContents
[F19:I19].ClearContents
[C19:E19].ClearContents
ActiveSheet.Shapes("Check Box 33").OLEFormat.Object.Value = 0
ActiveSheet.Shapes("Check Box 31").OLEFormat.Object.Value = 0
ActiveSheet.Shapes("Check Box 3").OLEFormat.Object.Value = 0
ActiveSheet.Shapes("Check Box 43").OLEFormat.Object.Value = 0
ActiveSheet.Shapes("Check Box 13").OLEFormat.Object.Value = 0
ActiveSheet.Shapes("Check Box 11").OLEFormat.Object.Value = 0
ActiveSheet.Shapes("Check Box 12").OLEFormat.Object.Value = 0
[C26:F27].ClearContents
[C29:F30].ClearContents
[G26:I27].ClearContents
[G29:I30].ClearContents
ActiveSheet.Shapes("Check Box 23").OLEFormat.Object.Value = 0
ActiveSheet.Shapes("Check Box 15").OLEFormat.Object.Value = 0
ActiveSheet.Shapes("Check Box 17").OLEFormat.Object.Value = 0
ActiveSheet.Shapes("Check Box 18").OLEFormat.Object.Value = 0
[C45:G45] = "5. Name: "
[C46:G46] = "Address: "
[H45:I45] = "SSN: "
ActiveSheet.Shapes("Check Box 21").OLEFormat.Object.Value = 0

Range("A19").Select
End Sub

Sub MA_Comp_Transfer_SectionA()
'
'Transfer Wages to Section A
Dim linea, lineb, linec As String

linea = [A71]
lineb = [E71]
linec = [I71]

msg = "Any information previously entered into 'Section A' will be cleared and updated with the new information from this section.  Are you sure you want to proceed?"
Ans = MsgBox(msg, vbYesNo)
If Ans = vbNo Then
    Exit Sub
End If

[A15:I19].ClearContents

'First Member
If [A71] <> "LN #" Then
[A15].Select
    Do Until ActiveCell.Value = Empty
        ActiveCell.Offset(1, 0).Select
    Loop
ActiveCell.Value = Right(linea, 2) & "/311"
ActiveCell.Offset(0, 2).Value = [D81]
[D13] = [A73]

'Puts values in for entire administrative period.
If [D90] <> 0 Then
    ActiveCell.Offset(0, 3).Value = [D90]
    [E13] = [A82]
    ActiveCell.Offset(0, 4).Value = [D99]
    [F13] = [A91]
End If
'Puts values in for entire 6-month period.
If [D108] <> 0 Or [D117] <> 0 Or [D126] <> 0 Then
    ActiveCell.Offset(0, 5).Value = [D108]
    [G13] = [A100]
    ActiveCell.Offset(0, 6).Value = [D117]
    [H13] = [A109]
    ActiveCell.Offset(0, 7).Value = [D126]
    [I13] = [A118]
End If
End If

'Second Member
If [E71] <> "LN #" Then
[A15].Select
    Do Until ActiveCell.Value = Empty
        ActiveCell.Offset(1, 0).Select
    Loop
ActiveCell.Value = Right(lineb, 2) & "/311"
ActiveCell.Offset(0, 2).Value = [H81]
If [H90] <> 0 Then
    ActiveCell.Offset(0, 3).Value = [H90]
    [E13] = [E82]
    ActiveCell.Offset(0, 4).Value = [H99]
    [F13] = [E91]
End If
If [H108] <> 0 Or [H117] <> 0 Or [H126] <> 0 Then
    ActiveCell.Offset(0, 5).Value = [H108]
    [G13] = [E100]
    ActiveCell.Offset(0, 6).Value = [H117]
    [H13] = [E109]
    ActiveCell.Offset(0, 7).Value = [H126]
    [I13] = [E118]
End If
End If

'Third Member
If [I71] <> "LN #" Then
[A15].Select
    Do Until ActiveCell.Value = Empty
        ActiveCell.Offset(1, 0).Select
    Loop
ActiveCell.Value = Right(linec, 2) & "/311"
ActiveCell.Offset(0, 2).Value = [L81]
If [L90] <> 0 Then
    ActiveCell.Offset(0, 3).Value = [L90]
    [E13] = [I82]
    ActiveCell.Offset(0, 4).Value = [L99]
    [F13] = [I91]
End If
If [L108] <> 0 Or [L117] <> 0 Or [L126] <> 0 Then
    ActiveCell.Offset(0, 5).Value = [L108]
    [G13] = [I100]
    ActiveCell.Offset(0, 6).Value = [L117]
    [H13] = [I109]
    ActiveCell.Offset(0, 7).Value = [L126]
    [I13] = [I118]
End If
End If

'Enters zeros into the blank cells in the administrative period if any member has income in the admin period
If [E15] <> Empty Or [F15] <> Empty Or [E16] <> Empty Or [F16] <> Empty Or [E17] <> Empty Or [F17] <> Empty Then
    If [D15] <> Empty Then
    [E15] = [D90]
    [F15] = [D99]
    End If
    If [D16] <> Empty Then
    [E16] = [H90]
    [F16] = [H99]
    End If
    If [D17] <> Empty Then
    [E17] = [L90]
    [F17] = [L99]
    End If
End If

'Enters zeros into the blank cells in the 6-month period if any member has income in the 6-month period
If [G15] <> Empty Or [H15] <> Empty Or [I15] <> Empty Or [G16] <> Empty Or [H16] <> Empty Or [I16] <> Empty Or [G17] <> Empty Or [H17] <> Empty Or [I17] <> Empty Then
    If [D15] <> Empty Or [D15] = 0 Then
    [G15] = [D108]
    [H15] = [D117]
    [I15] = [D126]
    End If
    If [D16] <> Empty Or [D16] = 0 Then
    [G16] = [H108]
    [H16] = [H117]
    [I16] = [H126]
    End If
    If [D17] <> Empty Or [D17] = 0 Then
    [G17] = [L108]
    [H17] = [L117]
    [I17] = [L126]
    End If
End If
End Sub

Sub MA_Comp_Clear_Wages()
'
'Clear Wages worksheet
msg = "Any information previously entered into this section will be deleted.  Are you sure you want to proceed?"
Ans = MsgBox(msg, vbYesNo)
If Ans = vbNo Then
    Exit Sub
End If

[A75:L79].ClearContents
[A84:L88].ClearContents
[A93:L97].ClearContents
[A102:L106].ClearContents
[A111:L115].ClearContents
[A120:L124].ClearContents

[A71] = "LN #"
[E71] = "LN #"
[I71] = "LN #"

[A72] = "Employer 1"
[C72] = "Employer 2"
[E72] = "Employer 1"
[G72] = "Employer 2"
[I72] = "Employer 1"
[K72] = "Employer 2"

End Sub

Sub MA_Comp_Clear_Income()
'
'Clear Income Comp Sheet
msg = "Any information previously entered into this worksheet will be deleted.  Are you sure you want to proceed?"
Ans = MsgBox(msg, vbYesNo)
If Ans = vbNo Then
    Exit Sub
End If

[A15:L19].ClearContents
[P25:V25].ClearContents
[P16:V16].ClearContents
[C24:L35].ClearContents
[C40:L45].ClearContents
[C52:L55].ClearContents
[C46:L46].ClearContents
[D13:L13].ClearContents
End Sub

Sub MA_Comp_Transfer_Workbook()
'
'
'Transfer to Workbook

Dim finalcol As String

finalcol = InputBox("Which column do you want to be transfered into Elements 371/372?" & vbNewLine & "Example: Q, R, or S?", "Final")
If finalcol = "" Then
    Exit Sub
End If
finalcol = UCase(finalcol)

If finalcol = "Q" Or finalcol = "R" Or finalcol = "S" Then
    Sheets("MA Workbook").[AB1425] = Sheets("MA Income Comp").Range(finalcol & "17").Value
    Sheets("MA Workbook").[AB1430] = Sheets("MA Income Comp").Range(finalcol & "19").Value
    Sheets("MA Workbook").[AB1432] = Sheets("MA Income Comp").Range(finalcol & "22").Value
    Sheets("MA Workbook").[AB1434] = Sheets("MA Income Comp").Range(finalcol & "24").Value
    Sheets("MA Workbook").Shapes("Check Box 570").OLEFormat.Object.Value = 1
    Sheets("MA Workbook").Shapes("Check Box 571").OLEFormat.Object.Value = 1
    Sheets("MA Workbook").Shapes("Check Box 572").OLEFormat.Object.Value = 1
    Sheets("MA Workbook").Shapes("Check Box 669").OLEFormat.Object.Value = 1
    Sheets("MA Workbook").Shapes("Check Box 671").OLEFormat.Object.Value = 1
Else
    MsgBox ("You did not enter a Q or R or S. Please try again")
    Exit Sub
End If

'If UCase(finalcol) = "Q" Then
'Sheets("MA Workbook").[AB1425] = Sheets("MA Income Comp").[Q17]
'Sheets("MA Workbook").[AB1430] = Sheets("MA Income Comp").[Q19]
'Sheets("MA Workbook").[AB1432] = Sheets("MA Income Comp").[Q22]
'Sheets("MA Workbook").[AB1434] = Sheets("MA Income Comp").[Q24]
'Sheets("MA Workbook").Shapes("Check Box 570").OLEFormat.Object.Value = 1
'Sheets("MA Workbook").Shapes("Check Box 571").OLEFormat.Object.Value = 1
'Sheets("MA Workbook").Shapes("Check Box 572").OLEFormat.Object.Value = 1
'Sheets("MA Workbook").Shapes("Check Box 669").OLEFormat.Object.Value = 1
'Sheets("MA Workbook").Shapes("Check Box 671").OLEFormat.Object.Value = 1
'End If

'If UCase(finalcol) = "R" Then
'Sheets("MA Workbook").[AB1425] = Sheets("MA Income Comp").[R17]
'Sheets("MA Workbook").[AB1430] = Sheets("MA Income Comp").[R19]
'Sheets("MA Workbook").[AB1432] = Sheets("MA Income Comp").[R22]
'Sheets("MA Workbook").[AB1434] = Sheets("MA Income Comp").[R24]
'Sheets("MA Workbook").Shapes("Check Box 570").OLEFormat.Object.Value = 1
'Sheets("MA Workbook").Shapes("Check Box 571").OLEFormat.Object.Value = 1
'Sheets("MA Workbook").Shapes("Check Box 572").OLEFormat.Object.Value = 1
'Sheets("MA Workbook").Shapes("Check Box 669").OLEFormat.Object.Value = 1
'Sheets("MA Workbook").Shapes("Check Box 671").OLEFormat.Object.Value = 1
'End If

'If UCase(finalcol) = "S" Then
'Sheets("MA Workbook").[AB1425] = Sheets("MA Income Comp").[S17]
'Sheets("MA Workbook").[AB1430] = Sheets("MA Income Comp").[S19]
'Sheets("MA Workbook").[AB1432] = Sheets("MA Income Comp").[S22]
'Sheets("MA Workbook").[AB1434] = Sheets("MA Income Comp").[S24]
'Sheets("MA Workbook").Shapes("Check Box 570").OLEFormat.Object.Value = 1
'Sheets("MA Workbook").Shapes("Check Box 571").OLEFormat.Object.Value = 1
'Sheets("MA Workbook").Shapes("Check Box 572").OLEFormat.Object.Value = 1
'Sheets("MA Workbook").Shapes("Check Box 669").OLEFormat.Object.Value = 1
'Sheets("MA Workbook").Shapes("Check Box 671").OLEFormat.Object.Value = 1
'End If

'QC Column
If Sheets("MA Workbook").Shapes("Check Box 565").OLEFormat.Object.Value <> 1 Then
Sheets("MA Workbook").[M1425] = Sheets("MA Income Comp").[P17]
Sheets("MA Workbook").[M1428] = Sheets("MA Income Comp").[P19]
Sheets("MA Workbook").[M1430] = Sheets("MA Income Comp").[P22]
Sheets("MA Workbook").[M1432] = Sheets("MA Income Comp").[P24]
Sheets("MA Workbook").Shapes("Check Box 566").OLEFormat.Object.Value = 1
Sheets("MA Workbook").Shapes("Check Box 567").OLEFormat.Object.Value = 1
Sheets("MA Workbook").Shapes("Check Box 568").OLEFormat.Object.Value = 1
Sheets("MA Workbook").Shapes("Check Box 670").OLEFormat.Object.Value = 1
Sheets("MA Workbook").Shapes("Check Box 672").OLEFormat.Object.Value = 1
End If

End Sub

'Sub PA721_Fill()
'
' PA721autofill Macro
'Takes information from the facesheet of the Workbook (looks for something entered in eligibility)
'
'
'This processing is done in populate when the worksheet is copied
'Dim EOMONTH As Date
'[P4] = Evaluate("=EOMONTH(AL4, 0)")
'[P3] = Evaluate("=EOMONTH(AL4, -1)+1")
'[P3].NumberFormat = "m/dd/yyyy"
'[P4].NumberFormat = "m/dd/yyyy"
'[Y3].NumberFormat = "m/dd/yyyy"
'[Y4].NumberFormat = "m/dd/yyyy"

'For i = 11 To 19
'    [D19].Select
'
'    Do Until ActiveCell.value = Empty
'        ActiveCell.Offset(1, 0).Select
'    Loop
'
'    If Sheets("MA Workbook").Cells(i, 37).value <> Empty Then
'        ActiveCell.Offset(0, -3).value = Sheets("MA Workbook").Cells(i, 10).value
'        ActiveCell.value = Sheets("MA Workbook").Cells(i, 12).value
'    End If
'Next i

'already done in populate
'[A33] = [AG3]
'Cat = [AG3]
'If Cat = "PAN-00" Or Cat = "PAN-80" Or Cat = "PJN-00" Or Cat = "PJN-80" Or Cat = "PAN-66" Or Cat = "PJN-66" Then
'    ActiveSheet.Shapes("Check Box 8").OLEFormat.Object.Value = 1
'Else
'    ActiveSheet.Shapes("Check Box 9").OLEFormat.Object.Value = 1
'End If

'End Sub
