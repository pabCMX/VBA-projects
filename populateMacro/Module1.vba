Attribute VB_Name = "Module1"
Sub tracking_supr()
    Range("M13") = Environ("USERNAME") & " " & Date
End Sub
Sub UserNameWindows()
'Find path to Examiner's files on this PC
Dim thiswb As Workbook

Set thiswb = ActiveWorkbook
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
    MsgBox "Network Drive to Examiner Files are NOT correct" & Chr(13) & _
        "Contact Valerie or Nicole"
    End
End If

pathdir = DLetter & "\SE Clerical Schedules\"
pathdir2 = DLetter & "\HBG Clerical Schedules\"
pathdir3 = DLetter & "\Pgh Clerical Schedules\"

If Len(Dir(pathdir, vbDirectory)) = 0 Then
    MsgBox "Path to Examiner's File: " & pathdir & " does NOT exists!!" & Chr(13) & _
        "Contact Valerie or Nicole"
    End
End If

Select Case Left(ActiveSheet.Name, 1)
    Case "1"  'TANF
        Range("AL5") = Environ("USERNAME") & " " & Date
        thiswb.Worksheets("TANF Workbook").Range("G41") = Environ("USERNAME")
        thiswb.Worksheets("TANF Workbook").Range("G44") = Date
        foldertxt = ""
            If Range("AI10") > 1 Then
                foldertxt = "Drop"
            ElseIf Range("AL10") = 1 Then
                foldertxt = "Clean"
            Else
                foldertxt = "Error"
            End If
            
        ActiveWorkbook.SAVE
        examnum = Val(Range("AO3") & Range("AP3"))
        
        If Left(Range("A10"), 1) = 1 Then  'TANF
                    If examnum = 98 Then
                         ActiveWorkbook.SaveAs fileName:=pathdir2 & "TANF\" & foldertxt & "\" & ActiveWorkbook.Name
                    Else
                        ActiveWorkbook.SaveAs fileName:=pathdir2 & "TANF\" & foldertxt & "\" & ActiveWorkbook.Name
                    End If
        
        Else  'GA
                    If examnum = 98 Then
                         ActiveWorkbook.SaveAs fileName:=pathdir & "GA\" & foldertxt & "\" & ActiveWorkbook.Name
                    Else
                        ActiveWorkbook.SaveAs fileName:=pathdir & "GA\" & foldertxt & "\" & ActiveWorkbook.Name
                    End If
        End If
    'PE
   ' If Left(thisws.Name, 2) = "24" Then
   '
   '     Range("AB3") = Environ("USERNAME") & " " & Date
   '     ' Set drop variable
   '      '   foldertxt = ""
   '      '   If Range("F16") <> 1 Then
   '      '       foldertxt = "Drop"
   '      '   ElseIf Range("F16") = 1 And Val(Range("S16")) = 1 Then
   '      '       foldertxt = "Clean"
    '     '   Else
    '     '       foldertxt = "Error"
    '     '   End If
    ''
    '        ActiveWorkbook.Save
    '
    '            examnum = Val(Range("AB11") & Range("AC11"))
    '                If examnum = 95 Then
     '                    ActiveWorkbook.SaveAs Filename:=pathdir2 & "PE\" & ActiveWorkbook.Name
     '
     ''              End If
    
 Case "9"  ' GA
        Range("AL5") = Environ("USERNAME") & " " & Date
        thiswb.Worksheets("GA Workbook").Range("G41") = Environ("USERNAME")
        thiswb.Worksheets("GA Workbook").Range("G44") = Date
        foldertxt = ""
            If Range("AI10") > 1 Then
                foldertxt = "Drop"
            ElseIf Range("AL10") = 1 Then
                foldertxt = "Clean"
            Else
                foldertxt = "Error"
            End If
            
        ActiveWorkbook.SAVE
        examnum = Val(Range("AO3") & Range("AP3"))
        
        If Left(Range("A10"), 1) = 1 Then  'TANF
                    If examnum = 98 Then
                         ActiveWorkbook.SaveAs fileName:=pathdir & "TANF\" & foldertxt & "\" & ActiveWorkbook.Name
                    Else
                        ActiveWorkbook.SaveAs fileName:=pathdir & "TANF\" & foldertxt & "\" & ActiveWorkbook.Name
                    End If
        
        Else  'GA
                    If examnum = 98 Then
                         ActiveWorkbook.SaveAs fileName:=pathdir & "GA\" & foldertxt & "\" & ActiveWorkbook.Name
                    Else
                        ActiveWorkbook.SaveAs fileName:=pathdir & "GA\" & foldertxt & "\" & ActiveWorkbook.Name
                    End If
        End If
    Case "2"  'MA Positive
        Range("AL5") = Environ("USERNAME") & " " & Date
        thiswb.Worksheets("MA Workbook").Range("F43") = Environ("USERNAME")
        thiswb.Worksheets("MA Workbook").Range("F45") = Date
        
        ' Set drop variable
            foldertxt = ""
            If Range("F16") <> 1 Then
                foldertxt = "Drop"
            ElseIf Range("F16") = 1 And Val(Range("S16")) = 1 Then
                foldertxt = "Clean"
            Else
                foldertxt = "Error"
            End If
            
            ActiveWorkbook.SAVE
        
                examnum = Val(Range("AO3") & Range("AP3"))
                    If examnum = 44 Or examnum = 38 Then
                         ActiveWorkbook.SaveAs fileName:=pathdir2 & "MA Positive\" & foldertxt & "\" & ActiveWorkbook.Name
                    'ElseIf examnum = 87 Or examnum = 45 Then
                    '     ActiveWorkbook.SaveAs Filename:=pathdir3 & "MA Positive\" & foldertxt & "\" & ActiveWorkbook.Name
                    Else
                        ActiveWorkbook.SaveAs fileName:=pathdir2 & "MA Positive\" & foldertxt & "\" & ActiveWorkbook.Name
                    End If
            
    
    Case "5"  'SNAP Positive
        'check snap expedited block in section 7
        If Range("B157") = "" Then
            Range("B157").Select
            MsgBox "The SNAP Expedited block (row 3, 1st block) in Section 7 must be filled in. Please fill in the block and click the Supervisor Approval button again."
        Else
            Range("AH2") = Environ("USERNAME") & " " & Date
            thiswb.Worksheets("FS Workbook").Range("G42") = Environ("USERNAME")
            thiswb.Worksheets("FS Workbook").Range("G44") = Date
            
            ' Set drop variable
            foldertxt = ""
            If Range("C22") > 1 Then
                foldertxt = "Drop"
            ElseIf Range("K22") = 1 Then
                foldertxt = "Clean"
            Else
                foldertxt = "Error"
            End If
            
            ActiveWorkbook.SAVE
                        
        examnum = Val(Range("AJ5") & Range("AK5"))
        If examnum = 53 Or examnum = 91 Then
            ActiveWorkbook.SaveAs fileName:=pathdir2 & "SNAP Positive\" & foldertxt & "\" & ActiveWorkbook.Name
        ElseIf examnum = 42 Or examnum = 46 Or examnum = 92 Or examnum = 40 Then
            ActiveWorkbook.SaveAs fileName:=pathdir2 & "SNAP Positive\" & foldertxt & "\" & ActiveWorkbook.Name
        Else
            ActiveWorkbook.SaveAs fileName:=pathdir2 & "SNAP Positive\" & foldertxt & "\" & ActiveWorkbook.Name
        End If
        
        'Timeliness Memo Reminder to Supervisors
        If Range("C149") = 2 Or Range("k149") = 11 Or Range("k149") = 12 Or Range("k149") = 13 Then
            MsgBox "Check for SNAP Timeliness Findings Memo", vbCritical, "Timeliness"
        End If
        
        If Range("k149") = 24 Or Range("k149") = 25 Or Range("k149") = 26 Or Range("k149") = 27 Then
            MsgBox "Check for SNAP Timeliness Info Memo", vbCritical, "Timeliness"
        End If
        
        End If
        
    Case "8"  'MA Negative
        Range("AB2") = Environ("USERNAME") & " " & Date
        foldertxt = ""
            If Range("M56") <> 0 Then
                foldertxt = "Drop"
            ElseIf (Range("C25") > 1 And Range("C25") < 5) Or Range("C40") = 2 Or _
                    ((Range("C40") > 2 And Range("C40") < 6) And Range("C56") = 2) Then
                foldertxt = "Error"
            Else
                foldertxt = "Clean"
            End If
            
            ActiveWorkbook.SAVE
            
        examnum = Val(Range("AB11") & Range("AC11"))
        If examnum = 48 Or examnum = 44 Or examnum = 38 Then
            ActiveWorkbook.SaveAs fileName:=pathdir2 & "MA Negative\" & foldertxt & "\" & ActiveWorkbook.Name
        'ElseIf examnum = 87 Or examnum = 45 Then
        '    ActiveWorkbook.SaveAs Filename:=pathdir3 & "MA Negative\" & foldertxt & "\" & ActiveWorkbook.Name
        Else
            ActiveWorkbook.SaveAs fileName:=pathdir2 & "MA Negative\" & foldertxt & "\" & ActiveWorkbook.Name
        End If
        
    Case "6"  'SNAP Negative
        Range("AC17") = Environ("USERNAME") & " " & Date
        tempArray = Split(Date, "/")
        
        Range("AA16") = tempArray(LBound(tempArray))
        Range("AD16") = tempArray(LBound(tempArray) + 1)
        Range("AG16") = tempArray(UBound(tempArray))
        
    
        foldertxt = ""
            'If Range("AF34") > 1 Then  'old code
            If Range("F29") > 1 Then 'new code - 11/2014
                foldertxt = "Drop"
            'ElseIf Range("E41") <> "" Then 'old code
            ElseIf Range("E47") <> "" Then 'new code - 11/2014
                foldertxt = "Error"
            Else
                foldertxt = "Clean"
            End If
            
            ActiveWorkbook.SAVE
                  
        examnum = Val(Range("W17"))
        If examnum = 53 Or examnum = 91 Then
            ActiveWorkbook.SaveAs fileName:=pathdir2 & "SNAP Negative\" & foldertxt & "\" & ActiveWorkbook.Name
        ElseIf examnum = 42 Or examnum = 46 Or examnum = 92 Or examnum = 40 Then
            ActiveWorkbook.SaveAs fileName:=pathdir2 & "SNAP Negative\" & foldertxt & "\" & ActiveWorkbook.Name
        Else
            ActiveWorkbook.SaveAs fileName:=pathdir2 & "SNAP Negative\" & foldertxt & "\" & ActiveWorkbook.Name
        End If
    
    End Select
    
    
    Range("A1").Select
    
        
End Sub
Sub ClericalApproval()
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
    MsgBox "Network Drive to Examiner Files are NOT correct" & Chr(13) & _
        "Contact Valerie or Nicole"
    End
End If

pathdir = DLetter & "\QCMIS Schedules\"

If Len(Dir(pathdir, vbDirectory)) = 0 Then
    MsgBox "Path to Examiner's File: " & pathdir & " does NOT exists!!" & Chr(13) & _
        "Contact Valerie or Nicole"
    End
End If

Select Case Left(ActiveSheet.Name, 1)
    Case "1", "2", "9"
        Range("AL6") = Environ("USERNAME") & " " & Date
    Case "5"
        Range("AH13") = Environ("USERNAME") & " " & Date
    Case "8"
        Range("AB3") = Environ("USERNAME") & " " & Date
    Case "6"
        Range("AK15") = Environ("USERNAME") & " " & Date
End Select

ActiveWorkbook.SAVE

'progadd = ""
'temparray = Split(ActiveWorkbook.Path, "\")
'For i = LBound(temparray) To UBound(temparray)
'    If InStr(temparray(i), "Clerical Schedules") > 0 Then
'        progadd = temparray(i + 1)
'        foldertxt = temparray(i + 2)
'        Exit For
'    End If
'Next i

'If progadd = "" Then
'    ActiveWorkbook.SaveAs Filename:=pathdir & ActiveWorkbook.Name
'Else
'    ActiveWorkbook.SaveAs Filename:=pathdir & progadd & "\" & foldertxt & "\" & ActiveWorkbook.Name
'End If

Select Case Left(ActiveSheet.Name, 1)
    Case "1", "9"
            Range("AL6") = Environ("USERNAME") & " " & Date
           foldertxt = ""
            If Range("AI10") > 1 Then
                foldertxt = "Drop"
            ElseIf Range("AL10") = 1 Then
                foldertxt = "Clean"
            Else
                foldertxt = "Error"
            End If
            
            ActiveWorkbook.SAVE

        If Left(Range("A10"), 1) = 1 Then
            ActiveWorkbook.SaveAs fileName:=pathdir & "TANF\" & foldertxt & "\" & ActiveWorkbook.Name
        Else
            ActiveWorkbook.SaveAs fileName:=pathdir & "GA\" & foldertxt & "\" & ActiveWorkbook.Name
        End If

    Case "2"
            Range("AL6") = Environ("USERNAME") & " " & Date
           foldertxt = ""
            If Range("F16") > 1 Then
                foldertxt = "Drop"
            ElseIf Range("F16") = 1 And Val(Range("S16")) = 1 Then
                foldertxt = "Clean"
            Else
                foldertxt = "Error"
            End If
            
            ActiveWorkbook.SAVE

            ActiveWorkbook.SaveAs fileName:=pathdir & "MA Positive\" & foldertxt & "\" & ActiveWorkbook.Name
       
    Case "5"
        Range("AH13") = Environ("USERNAME") & " " & Date
        ' Set drop variable
            foldertxt = ""
            If Range("C22") > 1 Then
                foldertxt = "Drop"
            ElseIf Range("K22") = 1 Then
                foldertxt = "Clean"
            Else
                foldertxt = "Error"
            End If

            ActiveWorkbook.SAVE

        ActiveWorkbook.SaveAs fileName:=pathdir & "SNAP Positive\" & foldertxt & "\" & ActiveWorkbook.Name

    Case "8"
        Range("AB3") = Environ("USERNAME") & " " & Date
        foldertxt = ""
            If Range("M56") <> 0 Then
                foldertxt = "Drop"
            ElseIf UCase(Range("AC1")) = UCase("Error") Then
                foldertxt = "Error"
            Else
                foldertxt = "Clean"
            End If

            ActiveWorkbook.SAVE

        ActiveWorkbook.SaveAs fileName:=pathdir & "MA Negative\" & foldertxt & "\" & ActiveWorkbook.Name

    Case "6"
        Range("AK15") = Environ("USERNAME")
        tempArray = Split(Date, "/")

        Range("AA16") = tempArray(LBound(tempArray))
        Range("AD16") = tempArray(LBound(tempArray) + 1)
        Range("AG16") = tempArray(UBound(tempArray))


        foldertxt = ""
            'If Range("AF34") > 1 Then 'old code
            If Range("F29") > 1 Then 'new code - 11/2014
                foldertxt = "Drop"
            'ElseIf Range("E41") <> "" Then 'old code
            ElseIf Range("E47") <> "" Then 'new code - 11/2014
                foldertxt = "Error"
            Else
                foldertxt = "Clean"
            End If

            ActiveWorkbook.SAVE

        ActiveWorkbook.SaveAs fileName:=pathdir & "SNAP Negative\" & foldertxt & "\" & ActiveWorkbook.Name
    End Select

    
    Range("A1").Select
    
    
End Sub
Sub SuprWorkbook()

    Select Case ActiveSheet.Name
    Case "TANF Workbook", "GA Workbook"
        Range("G41") = Environ("USERNAME")
        Range("G44") = Date
    Case "FS Workbook"
        Range("G42") = Environ("USERNAME")
        Range("G44") = Date
    End Select
    
    Range("A1").Select
        
End Sub
Sub Macro2()

    'Cells(3, 8) = 0
   
    Range("A:H,AF:AK").Select
    
    Sheets("FS Computation").Unprotect "QC"

    'clearing any contents with numbers that were keyed in
    Selection.Range("A2:H67,AF7:AK17").SpecialCells(xlCellTypeConstants, 1).Select
    Selection.ClearContents
    
    'clearing line number and income type
    Range("A8:A10,A18:A22").Select
    Selection.ClearContents
    
    For i = 2 To 8
        Cells(44, i) = "No Utilities"
    Next i
    
    
    For i = 2 To 8
        If Cells(45, i) = "" Then
            'IF(B$11="","",IF($H$1>40086,VLOOKUP(B$44,$AB$30:$AC$34,2,FALSE),VLOOKUP(B$44,$AB$24:$AC$28,2,FALSE)))
            TempStr = "=IF(R11C" & i & "="""","""",IF(R1C8>40086,VLOOKUP(R44C" & i & ",R30C28:R34C29,2,FALSE),VLOOKUP(R44C" & i & ",R24C28:R28C29,2,FALSE)))"
            Cells(45, i).FormulaR1C1 = TempStr
        End If
    Next i
        
    Sheets("FS Computation").Protect "QC"
   
    Range("A1").Select

End Sub

Sub finalresults()
If Range("E62") = Range("D62") Then
        Range("D8:D11").Select
        Selection.Copy
        Range("C8:C11").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Range("D13:D14").Select
        Selection.Copy
        Range("C13:C14").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Range("D18:D22").Select
        Selection.Copy
        Range("C18:C22").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        'Range("D26").Select
        'Selection.Copy
        'Range("C26").Select
        'Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        ':=False, Transpose:=False
        Range("D34").Select
        Selection.Copy
        Range("C34").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Range("D36").Select
        Selection.Copy
        Range("C36").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Range("D38").Select
        Selection.Copy
        Range("C38").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Range("D42:D47").Select
        Selection.Copy
        Range("C42:C47").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Range("D66").Select
        Selection.Copy
        Range("C66").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Range("D51:D52").Select
        Selection.Copy
        Range("C51:C52").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Range("D61").Select
        Selection.Copy
        Range("C61").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Application.CutCopyMode = False

ElseIf Range("E11") = "" Then
        Range("D8:D11").Select
        Selection.Copy
        Range("C8:C11").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Range("D13:D14").Select
        Selection.Copy
        Range("C13:C14").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Range("D18:D22").Select
        Selection.Copy
        Range("C18:C22").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        'Range("D26").Select
        'Selection.Copy
        'Range("C26").Select
        'Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        ':=False, Transpose:=False
        Range("D34").Select
        Selection.Copy
        Range("C34").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Range("D36").Select
        Selection.Copy
        Range("C36").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Range("D38").Select
        Selection.Copy
        Range("C38").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Range("D42:D47").Select
        Selection.Copy
        Range("C42:C47").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Range("D66").Select
        Selection.Copy
        Range("C66").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Range("D51:D52").Select
        Selection.Copy
        Range("C51:C52").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Range("D61").Select
        Selection.Copy
        Range("C61").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Application.CutCopyMode = False
        
ElseIf Range("D11") = "" Then
         Range("E8:E11").Select
        Selection.Copy
        Range("C8:C11").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Range("E13:E14").Select
        Selection.Copy
        Range("C13:C14").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Range("E18:E22").Select
        Selection.Copy
        Range("C18:C22").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        'Range("E26").Select
        'Selection.Copy
        'Range("C26").Select
        'Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        ':=False, Transpose:=False
        Range("E34").Select
        Selection.Copy
        Range("C34").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Range("E36").Select
        Selection.Copy
        Range("C36").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Range("E38").Select
        Selection.Copy
        Range("C38").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Range("E42:E47").Select
        Selection.Copy
        Range("C42:C47").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Range("E66").Select
        Selection.Copy
        Range("C66").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Range("E51:E52").Select
        Selection.Copy
        Range("C51:C52").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Range("E61").Select
        Selection.Copy
        Range("C61").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Application.CutCopyMode = False


ElseIf Abs(Range("E62")) < Abs(Range("D62")) Then
        Range("E8:E11").Select
        Selection.Copy
        Range("C8:C11").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Range("E13:E14").Select
        Selection.Copy
        Range("C13:C14").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Range("E18:E22").Select
        Selection.Copy
        Range("C18:C22").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        'Range("E26").Select
        'Selection.Copy
        'Range("C26").Select
        'Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        ':=False, Transpose:=False
        Range("E34").Select
        Selection.Copy
        Range("C34").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Range("E36").Select
        Selection.Copy
        Range("C36").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Range("E38").Select
        Selection.Copy
        Range("C38").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Range("E42:E47").Select
        Selection.Copy
        Range("C42:C47").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Range("E66").Select
        Selection.Copy
        Range("C66").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Range("E51:E52").Select
        Selection.Copy
        Range("C51:C52").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Range("E61").Select
        Selection.Copy
        Range("C61").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Application.CutCopyMode = False
        
    Else
        Range("D8:D11").Select
        Selection.Copy
        Range("C8:C11").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Range("D13:D14").Select
        Selection.Copy
        Range("C13:C14").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Range("D18:D22").Select
        Selection.Copy
        Range("C18:C22").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        'Range("D26").Select
        'Selection.Copy
        'Range("C26").Select
        'Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        ':=False, Transpose:=False
        Range("D34").Select
        Selection.Copy
        Range("C34").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Range("D36").Select
        Selection.Copy
        Range("C36").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Range("D38").Select
        Selection.Copy
        Range("C38").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Range("D42:D47").Select
        Selection.Copy
        Range("C42:C47").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Range("D66").Select
        Selection.Copy
        Range("C66").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Range("D51:D52").Select
        Selection.Copy
        Range("C51:C52").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Range("D61").Select
        Selection.Copy
        Range("C61").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Application.CutCopyMode = False
    End If
   
    Dim TempStr As String, index As Integer
    Dim ws As Worksheet, thisws As Worksheet
    Dim linnum As String, inctype As String

    ' find name of schedule spreadsheet
    For Each ws In ThisWorkbook.Worksheets
    If Left(ws.Name, 2) = "50" Or Left(ws.Name, 2) = "51" Or Left(ws.Name, 2) = "55" Then
        Set thisws = ws
    End If
    Next ws
    
    ' blank out income amounts in schedule
    For j = 131 To 143 Step 3
        For k = 9 To 36 Step 9
            thisws.Cells(j, k) = ""
        Next k
    Next j

' loop through earned income line numbers and income types
    For i = 8 To 10
        If Cells(i, 3) <> "" Then
         If Cells(i, 3) <> 0 Then
            TempStr = Range("a" & i)
            If Mid(TempStr, 3, 1) <> "/" Then
                MsgBox "Please enter line number and income type in format ##/##"
                Exit Sub
            End If
            temparr = Split(TempStr, "/")
            index = LBound(temparr)
            linnum = temparr(index)
            inctype = temparr(index + 1)
            
            'checking whether the income type is a valid selection
            flag = 0
            For n = 126 To 147
                If thisws.Range("BB" & n) = inctype Then
                 flag = 1
                 Exit For
                End If
            Next n
            
            If flag = 0 Then
                MsgBox "Income type on row " & i & " is invalid.  Please enter a valid income type"
                Range("a" & i).Select
                Exit Sub
            End If
            
            ' loop through section 5 of schedule and put income amounts
            ' in rows with correct line #s and income types
            flag = 0
            For j = 131 To 143 Step 3
               If thisws.Cells(j, 2) = linnum Then
                    flag = 1
                   For k = 5 To 32 Step 9
                       If thisws.Cells(j, k) = inctype Then
                            totalamt = Cells(i, 3) + thisws.Cells(j, k + 4)
                            If totalamt > 0 Then
                                thisws.Cells(j, k + 4) = totalamt
                            Else
                                thisws.Cells(j, k + 4) = ""
                                thisws.Cells(j, k) = ""
                            End If
                            Exit For
                        'If blank is found, then create new income type entry
                        ElseIf thisws.Cells(j, k) = "" Then
                            thisws.Cells(j, k) = inctype
                            thisws.Cells(j, k + 4) = Cells(i, 3)
                            Exit For
                       End If
            
                    Next k
                End If
            Next j
            If flag = 0 Then
            For j = 131 To 143 Step 3
                If thisws.Cells(j, 2) = "" Then
                    thisws.Cells(j, 2) = linnum
                    thisws.Cells(j, 5) = inctype
                    thisws.Cells(j, 9) = Cells(i, 3)
                    Exit For
                End If
            Next j
            End If
         End If
        End If
    Next i
   
' loop through unearned income line numbers and income types
    For i = 18 To 22
        If Cells(i, 3) <> "" Then
        
            TempStr = Range("a" & i)
            If Mid(TempStr, 3, 1) <> "/" Then
                MsgBox "Please enter line number and income type in format ##/##"
                Exit Sub
            End If
            temparr = Split(TempStr, "/")
            index = LBound(temparr)
            linnum = temparr(index)
            inctype = temparr(index + 1)
             
             'checking whether the income type is a valid selection
            flag = 0
            For n = 126 To 147
                If thisws.Range("BB" & n) = inctype Then
                 flag = 1
                 Exit For
                End If
            Next n
            
            If flag = 0 Then
                MsgBox "Income type on row " & i & " is invalid.  Please enter a valid income type"
                Range("a" & i).Select
                Exit Sub
            End If
            ' loop through section 5 of schedule and put income amounts
            ' in rows with correct line #s and income types
            flag = 0
            For j = 131 To 143 Step 3
                If thisws.Cells(j, 2) = linnum Then
                    flag = 1
                    For k = 5 To 32 Step 9
                        If thisws.Cells(j, k) = inctype Then
                            totalamt = Cells(i, 3) + thisws.Cells(j, k + 4)
                            If totalamt > 0 Then
                                thisws.Cells(j, k + 4) = totalamt
                            Else
                                thisws.Cells(j, k + 4) = ""
                                thisws.Cells(j, k) = ""
                            End If
                            Exit For
                        'If blank is found, then create new income type entry
                        ElseIf thisws.Cells(j, k) = "" Then
                            thisws.Cells(j, k) = inctype
                            thisws.Cells(j, k + 4) = Cells(i, 3)
                            Exit For
                        End If
                            
                    Next k
                 End If
            Next j
            If flag = 0 Then
            For j = 131 To 143 Step 3
                If thisws.Cells(j, 2) = "" Then
                    thisws.Cells(j, 2) = linnum
                    thisws.Cells(j, 5) = inctype
                    thisws.Cells(j, 9) = Cells(i, 3)
                    Exit For
                End If
            Next j
            End If
        End If
    Next i
    
    Dim found As Integer
    'blank out line number if only line number is present
    For j = 131 To 143 Step 3
        If thisws.Cells(j, 2) <> "" Then
            found = 0
            For k = 5 To 32 Step 9
                If thisws.Cells(j, k) <> "" Then
                    found = 1
                    Exit For
                End If
            Next k
            If found = 0 Then thisws.Cells(j, 2) = ""
        End If
    Next j

    Range("C60").Select
    
    
End Sub
Sub SelfEmployment()

Range("AF7").Select
ActiveWindow.ScrollColumn = 20

End Sub
Sub moveselfemp()

Range("A2").Select

End Sub
Sub GAClear()

    'Cells(3, 8) = 0
   
    Range("A:L").Select
    
    Sheets("GA Computation").Unprotect "QC"
   
    'clearing any contents with numbers that were keyed in
    Selection.SpecialCells(xlCellTypeConstants, 1).Select
    Selection.ClearContents
  
    'clearing name in cell B1
    Range("B1").Select
    Selection.ClearContents
    
    'clearing name in cell C2
    Range("C2").Select
    Selection.ClearContents
    
    'clearing name in cell E2
    Range("E2").Select
    Selection.ClearContents
    
    Range("A7") = "1.  line              / "
    Range("A8") = "2.  line              / "
    Range("A9") = "3.  line              / "
    Range("A15") = "7.  line              / "
    Range("A16") = "8.  line              / "
    Range("A17") = "9.  line              / "
    Range("A18") = "10.  line              / "
    Range("B76") = "Comments:"
    
    For i = 3 To 9
     
        TempStr = "=IF(R52C" & i & ">R53C" & i & ",0,IF(AND(R69C" & i & "=" & """""" & ",R70C" & i & "=" & """""" & ")," & """""" & ",IF(R69C" & i & "=" & """""" & ",R70C" & i & ",R69C" & i & ")))"
        Cells(71, i).FormulaR1C1 = TempStr
        
    Next i
    
    For i = 11 To 12
     
        TempStr = "=IF(R52C" & i & ">R53C" & i & ",0,IF(AND(R69C" & i & "=" & """""" & ",R70C" & i & "=" & """""" & ")," & """""" & ",IF(R69C" & i & "=" & """""" & ",R70C" & i & ",R69C" & i & ")))"
        Cells(71, i).FormulaR1C1 = TempStr
        
    Next i
    
        
    Sheets("GA Computation").Protect "QC"
   
    Range("A1").Select

End Sub
Sub GAcompsheet()

For i = 3 To 14
    If i <> 10 Then
If Cells(65, i) > 0 Then

    If Cells(63, i) < 101 Then
        Hundint = Application.WorksheetFunction.RoundDown((Cells(63, i) / 100) - 1, 0)
    Else
        Hundint = Application.WorksheetFunction.RoundDown((Cells(63, i) / 100), 0)
    End If

 Application.EnableEvents = False
 If Cells(65, i) > 14 Then
 MsgBox ("Proration tables only go up to 14 days.  Please ask your Supervisor how to prorate beyond 14 days.")
 Else
 Select Case Cells(65, i).Value
 
  Case 1
    Cells(68, i) = 6.6 * Hundint + Round(0.0000472884 + 0.065710554 * (Cells(63, i) - 100 * Hundint), 1)
  Case 2
    Cells(68, i) = 13.1 * Hundint + Round(-0.00017559 + 0.131430335 * (Cells(63, i) - 100 * Hundint), 1)
  Case 3
  If Cells(63, i) = 87.5 Then
    Cells(68, i) = 17.2
  Else
    Cells(68, i) = 19.7 * Hundint + Round(-0.00010194 + 0.197144297 * (Cells(63, i) - 100 * Hundint), 1)
  End If
  Case 4
  If Cells(63, i) = 5.9 Then
    Cells(68, i) = 1.5
  Else
    Cells(68, i) = 26.3 * Hundint + Round(-0.000130429 + 0.262858301 * (Cells(63, i) - 100 * Hundint), 1)
  End If
  Case 5
  If Cells(63, i) = 6.9 Then
    Cells(68, i) = 2.2
  Else
    Cells(68, i) = 32.9 * Hundint + Round(-0.001141619 + 0.328581779 * (Cells(63, i) - 100 * Hundint), 1)
  End If
  Case 6
  If Cells(63, i) = 100.3 Then
    Cells(68, i) = 39.6
  Else
    Cells(68, i) = 39.4 * Hundint + Round(-0.00018061 + 0.3942901 * (Cells(63, i) - 100 * Hundint), 1)
  End If
  Case 7
  If Cells(63, i) = 50.3 Then
    Cells(68, i) = 23.2
  Else
    Cells(68, i) = 46 * Hundint + Round(0.001345477 + 0.459995182 * (Cells(63, i) - 100 * Hundint), 1)
  End If
  Case 8
    Cells(68, i) = 52.6 * Hundint + Round(-0.0000486654 + 0.5257144 * (Cells(63, i) - 100 * Hundint), 1)
  Case 9
    Cells(68, i) = 59.1 * Hundint + Round(-0.000156025 + 0.591429951 * (Cells(63, i) - 100 * Hundint), 1)
  Case 10
    Cells(68, i) = 65.7 * Hundint + Round(0.000167077 + 0.657139017 * (Cells(63, i) - 100 * Hundint), 1)
  Case 11
    Cells(68, i) = 72.3 * Hundint + Round(-0.0000937807 + 0.722855619 * (Cells(63, i) - 100 * Hundint), 1)
  Case 12
  If Cells(63, i) = 46.6 Then
    Cells(68, i) = 36.8
  Else
    Cells(68, i) = 78.9 * Hundint + Round(0.0000808591 + 0.788573487 * (Cells(63, i) - 100 * Hundint), 1)
  End If
  Case 13
    Cells(68, i) = 85.4 * Hundint + Round(0.0000748837 + 0.85428116 * (Cells(63, i) - 100 * Hundint), 1)
  Case 14
    Cells(68, i) = 92 * Hundint + Round(0.000024456 + 0.91999952 * (Cells(63, i) - 100 * Hundint), 1)
  
  End Select
  End If
   Application.EnableEvents = True
  End If
  
  End If
  
  Next i
  
  End Sub
  
Sub Compute()

Range("K7").Select

End Sub
Sub GAshow_form1()
    GAUserForm1.Show
End Sub
Sub GAshow_form2()
    GAUserForm2.Show
End Sub
Sub redisplayGAform1()
    GAUserForm1.Show
End Sub
Sub redisplayGAform2()
    GAUserForm2.Show
End Sub

Sub GAfinalresults()

    columnletter = Right(Range("AL77"), 1)
    rownumber = 71
    'If columnletter = "B" Then rownumber = 71
    TempStr = columnletter & rownumber
    Range(TempStr) = Range("K71") + Range("L71")
    Range(TempStr).Select

End Sub
Sub GAfinaldetermination()

    clet = Right(Range("AL78"), 1)
    Range(clet & "7:" & clet & "11").Copy
    Range("C7").PasteSpecial
    Range(clet & "15:" & clet & "18").Copy
    Range("C15").PasteSpecial
    Range(clet & "21").Copy
    Range("C21").PasteSpecial
    Range(clet & "25").Copy
    Range("C25").PasteSpecial
    Range(clet & "27").Copy
    Range("C27").PasteSpecial
    Range(clet & "29").Copy
    Range("C29").PasteSpecial
    Range(clet & "32").Copy
    Range("C32").PasteSpecial
    Range(clet & "35").Copy
    Range("C35").PasteSpecial
    Range(clet & "49").Copy
    Range("C49").PasteSpecial
    Range(clet & "54").Copy
    Range("C54").PasteSpecial
    Range(clet & "65").Copy
    Range("C65").PasteSpecial
    Range(clet & "67").Copy
    Range("C67").PasteSpecial
    Range(clet & "68").Copy
    Range("C68").PasteSpecial
    Range(clet & "73").Copy
    Range("C73").PasteSpecial

End Sub



'Sub check_box_80_click()
 '       Dim ws As Worksheet, thisws As Worksheet

    ' find name of schedule spreadsheet
  '  For Each ws In ThisWorkbook.Worksheets
   ' If Left(ws.Name, 2) = "50" Or Left(ws.Name, 2) = "51" Or Left(ws.Name, 2) = "55" Then
    '    Set thisws = ws
     '   Exit For
   ' End If
   ' Next ws
        
'clickyes = 0
'If Sheets("FS Workbook").Shapes("CB 148").OLEFormat.Object.Value = 1 Then
 '   For i = 89 To 122 Step 3
  '      If thisws.Range("B" & i) <> "" Then
   '         thisws.Range("AH" & i) = 2
    '        clickyes = 1
     '   End If
    'Next i
'ElseIf Sheets("FS Workbook").Shapes("CB 148").OLEFormat.Object.Value = -4146 Then
 '   For i = 89 To 122 Step 3
  '      If thisws.Range("B" & i) <> "" Then
   ''         thisws.Range("AH" & i) = ""
     '   End If
   ' Next i
'End If
'If clickyes = 0 Then
'If Sheets("FS Workbook").Shapes("CB 1112").OLEFormat.Object.Value = 1 Then
'    For i = 89 To 122 Step 3
'        If thisws.Range("B" & i) <> "" Then
'            thisws.Range("AH" & i) = 1
'        End If
'    Next i
'ElseIf Sheets("FS Workbook").Shapes("CB 1112").OLEFormat.Object.Value = -4146 Then
'    For i = 89 To 122 Step 3
'        If thisws.Range("B" & i) <> "" Then
'            thisws.Range("AH" & i) = ""
'        End If
'    Next i
'End If
'End If

'End Sub

'Sub check_box_671_click()
'        Dim ws As Worksheet, thisws As Worksheet

    ' find name of schedule spreadsheet
'    For Each ws In ThisWorkbook.Worksheets
'    If Left(ws.Name, 2) = "50" Or Left(ws.Name, 2) = "51" Or Left(ws.Name, 2) = "55" Then
'        Set thisws = ws
'        Exit For
'    End If
'    Next ws
    
'If Sheets("FS Workbook").Shapes("CB 1112").OLEFormat.Object.Value = 1 Then
'    For i = 89 To 122 Step 3
'        If thisws.Range("B" & i) <> "" Then
'            thisws.Range("AH" & i) = 1
'        End If
'    Next i
'ElseIf Sheets("FS Workbook").Shapes("CB 1112").OLEFormat.Object.Value = -4146 Then
'    For i = 89 To 122 Step 3
'        If thisws.Range("B" & i) <> "" Then
'            thisws.Range("AH" & i) = ""
'        End If
'    Next i
'End If
'End Sub

