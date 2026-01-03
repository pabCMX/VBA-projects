Attribute VB_Name = "populate_mod"
Public wb_bis As Workbook, wb As Workbook
Public Sh As Worksheet
Dim maxrow As Integer
Dim revmonth As Long
Dim fileName As String
Dim TempStr As String
Public program As String
Dim DLetter As String, pathdir As String

Sub showpopulate_form50()
    UserForm50.Show
End Sub
Sub showpopulate_form3()
    UserForm3.Show
End Sub
Sub redisplayform3()
    UserForm3.Show
End Sub

Sub Review_Schedule()
Attribute Review_Schedule.VB_ProcData.VB_Invoke_Func = " \n14"
'
' addresses Macro
' Macro recorded 3/25/2008 by


Dim Start_Range As Integer
Dim End_Range As Integer
Dim monthstr As String
Dim TempStr As String
Dim input_file_bis As String, popwb As Workbook
Dim revmn As Date

' Find current directory
    
    strcurdir = ActiveWorkbook.Path & "\"

'Find path to Examiner's files on this PC
Set WshNetwork = CreateObject("WScript.Network")
Set oDrives = WshNetwork.EnumNetworkDrives
Dim DUNC As String

tempstr1 = oDrives.Count

DLetter = ""
For i = 0 To oDrives.Count - 1 Step 2
    DUNC = "" & oDrives.Item(i + 1) & ""
    If LCase(DUNC) = "\\hsedcprapfpp001\oim\pwimdaubts04\data\stat" Then
        DLetter = "" & oDrives.Item(i) & "\DQC"
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
        "Contact Valerie or Matt"
    End
End If

pathdir = DLetter & "\Schedules by Examiner Number\"

If Len(Dir(pathdir, vbDirectory)) = 0 Then
    MsgBox "Path to Examiner's File: " & pathdir & " does NOT exists!!" & Chr(13) & _
        "Contact Valerie or Nicole"
    End
End If

Set popwb = ActiveWorkbook

' Find file of records file to open
    'ChDrive "\\dhs\share\oim\pwimdaubts04\data"
    'ChDir "\stat\DQC"
    input_file = Application _
    .GetOpenFilename(" File of Records(*.xlsm), *.xlsm", , _
    "Select File of Record file")
    If input_file = "False" Then Exit Sub
    
    Application.ScreenUpdating = False

'Read which program and month from user formpop1 input
    program = Cells(7, 23)
    revmonth = Cells(7, 26)

    If program = "FS Positive" Or program = "FS Supplemental" Then
        ' Find BIS Delimited file to open
        input_file_bis = Application _
            .GetOpenFilename(" BIS SNAP Positive or Supplemental Delimited Excel File (*.xlsx), *.xlsx", , _
            "Select BIS SNAP Positive or Supplemental Delimited Excel file")
        If input_file_bis = "False" Then Exit Sub
    ElseIf program = "FS Negative" Then
        ' Find BIS Delimited file to open
        input_file_bis = Application _
            .GetOpenFilename(" BIS SNAP Negative Delimited Excel File (*.xlsx), *.xlsx", , _
            "Select BIS SNAP Negative Delimited Excel file")
        If input_file_bis = "False" Then Exit Sub
    ElseIf program = "TANF" Then
        ' Find BIS Delimited file to open
        input_file_bis = Application _
            .GetOpenFilename(" BIS TANF Delimited Excel File (*.xlsx), *.xlsx", , _
            "Select BIS TANF Delimited Excel file")
        If input_file_bis = "False" Then Exit Sub
'    ElseIf program = "MA Positive" Then
'        ' Find BIS Delimited file to open
'        input_file_bis = Application _
'            .GetOpenFilename(" BIS MA Positive Delimited Excel File (*.xlsx), *.xlsx", , _
'            "Select BIS MA Positive Delimited Excel file")
'        If input_file_bis = "False" Then Exit Sub
    End If
  
'create new file with program and month
    monthstr = Application.WorksheetFunction.Text(revmonth, "MMMM YYYY")
    fileName = "Review Schedule for " & program & " " & monthstr & ".xlsx"

Application.DisplayAlerts = False
    Workbooks.Add
          ActiveWorkbook.SaveAs fileName:= _
          fileName, FileFormat:=51, _
          Password:="", WriteResPassword:="", ReadOnlyRecommended:=False _
          , CreateBackup:=False
Application.DisplayAlerts = True

    If program = "FS Positive" Or program = "FS Supplemental" Then
        'Open BIS Delimited file
        Workbooks.Open fileName:=input_file_bis, UpdateLinks:=False
        Set wb_bis = ActiveWorkbook
    ElseIf program = "FS Negative" Then
        Workbooks.Open fileName:=input_file_bis, UpdateLinks:=False
        Set wb_bis = ActiveWorkbook
    ElseIf program = "TANF" Then
        Workbooks.Open fileName:=input_file_bis, UpdateLinks:=False
        Set wb_bis = ActiveWorkbook
'    ElseIf program = "MA Positive" Then
'        Workbooks.Open FileName:=input_file_bis, UpdateLinks:=False
'        Set wb_bis = ActiveWorkbook
    End If

   'Open file of records
    Workbooks.Open fileName:=input_file, UpdateLinks:=False
    
    If program = "TANF" Or program = "GA" Or program = "FS Supplemental" Or program = "FS Positive" Or program = "FS Negative" Then
        Sheets("FS Cash main file").Select
        Sheets("FS Cash main file").Copy before:=Workbooks(fileName).Sheets(2)
        Sheets("FS Cash main file").Name = "Temp"
    ElseIf program = "MA Closed" Or program = "MA CHIP Closed" Or program = "MA Positive" Or program = "MA Rejected" Or program = "MA CHIP Rejected" Or program = "MA PE" Then
        Sheets("MA main file").Select
        Sheets("MA main file").Copy before:=Workbooks(fileName).Sheets(2)
        Sheets("MA main file").Name = "Temp"
    Else
        Sheets("CAR main file").Select
        Sheets("CAR main file").Copy before:=Workbooks(fileName).Sheets(2)
        Sheets("CAR main file").Name = "Temp"
    End If
    
    'Setting explicit references to needed workbooks and worksheets before managing files
    Dim wbPopulate As Workbook
    Dim wsTemp As Worksheet
    Set wbPopulate = Workbooks(fileName)
    Set wsTemp = wbPopulate.Sheets("Temp")
    
    'Making sure any loading happens now before closing
    DoEvents

    
    ' Find input file name
    testarr = Split(input_file, "\")
    upper = UBound(testarr)
    file_name = testarr(upper)
    
Application.DisplayAlerts = False
        
     'Close file of records
     Windows(file_name).Close
     
Application.DisplayAlerts = True
   
    'Waiting for more events before reactivating, then waiting for activates
    DoEvents
    wbPopulate.Activate
    wsTemp.Select
    DoEvents
         
    ' Find maximum row in temp spreadsheet
    maxrow = wsTemp.Cells(wsTemp.Rows.Count, "B").End(xlUp).Row
     
    'test for month in file
    flag = 0
    For i = 2 To maxrow
       If wsTemp.Cells(i, 2) = revmonth Then
          flag = 1
          Exit For
      End If
    Next i

    If flag = 0 Then
      MsgBox "Month not Found"
      Exit Sub
    End If
    
    
    ' Sort by column B which contains review numbers
    Range("A2:II" & maxrow).Sort Key1:=Range("A2"), Order1:=xlAscending, Header:= _
        xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortNormal
        
'find what type of review schedule that was selected
   Select Case program
    Case "GA"
         Start_Range = 90
         End_Range = 90
   ' Case "MA PE"
   '      Start_Range = 24
   '      End_Range = 24
    Case "MA Positive"
         Start_Range = 20
         End_Range = 25
    Case "FS Positive"
         Start_Range = 50
         End_Range = 51
    Case "FS Supplemental"
         Start_Range = 55
         End_Range = 55
    Case "FS Negative"
         Start_Range = 60
         End_Range = 66
    Case "TANF"
         Start_Range = 14
         End_Range = 14
    Case "TANF CAR"
         Start_Range = 34
         End_Range = 34
    Case "MA Closed"
         Start_Range = 80
         End_Range = 80
    Case "MA CHIP Closed"
         Start_Range = 84
         End_Range = 84
    Case "MA Rejected"
         Start_Range = 81
         End_Range = 81
    Case "MA CHIP Rejected"
         Start_Range = 85
         End_Range = 85
    Case "MA Negative"
         Start_Range = 80
         End_Range = 83
    End Select
    
'delete unwanted cases and months
    For i = maxrow To 1 Step -1
      If Val(Left(Cells(i, 1), 2)) < Start_Range Or Val(Left(Cells(i, 1), 2)) > End_Range Or Cells(i, 2) <> revmonth Then
       Rows(i).Delete
     End If
    Next i

'determining if there are any schedules in temp file
    If Range("A1") = "" Then
        MsgBox ("No schedules found for selected program and month")
        Windows("populate.xlsm").Activate
        Workbooks(fileName).Close (False)
        End
    End If
      
     ' Find maximum row in temp spreadsheet
    maxrow = ActiveSheet.Cells.Find(What:="*", _
        SearchDirection:=xlPrevious, _
        SearchOrder:=xlByRows).Row
 
 
    'Windows("populate.xlsm").Activate
    Set wbfile = Workbooks(fileName)
    'Workbooks(Filename).Activate
'Copy Review Schedule over to the filename
    Select Case program
    Case "GA"
         'popwb.Sheets("GA").Select
         popwb.Sheets("GA").Copy before:=wbfile.Sheets(1)
    Case "MA Positive"
         'popwb.Sheets("MA Pos").Select
         popwb.Sheets("MA Pos").Copy before:=wbfile.Sheets(1)
    Case "FS Positive", "FS Supplemental"
         'popwb.Sheets("FS Pos").Select
         popwb.Sheets("FS Pos").Copy before:=wbfile.Sheets(1)
    Case "FS Negative"
         'popwb.Sheets("FS Neg").Select
         popwb.Sheets("FS Neg").Copy before:=wbfile.Sheets(1)
    Case "TANF"
         'popwb.Sheets("TANF").Select
         popwb.Sheets("TANF").Copy before:=wbfile.Sheets(1)
    Case "TANF CAR"
         'popwb.Sheets("TANF CAR").Select
         popwb.Sheets("TANF CAR").Copy before:=wbfile.Sheets(1)
    Case "MA Closed"
         'popwb.Sheets("MA Neg").Select
         popwb.Sheets("MA Neg").Copy before:=wbfile.Sheets(1)
    Case "MA CHIP Closed"
         'popwb.Sheets("MA Neg").Select
         popwb.Sheets("MA Neg").Copy before:=wbfile.Sheets(1)
    Case "MA Rejected"
         'popwb.Sheets("MA Neg").Select
         popwb.Sheets("MA Neg").Copy before:=wbfile.Sheets(1)
    Case "MA CHIP Rejected"
         'popwb.Sheets("MA Neg").Select
         popwb.Sheets("MA Neg").Copy before:=wbfile.Sheets(1)
    Case "MA PE"
         'popwb.Sheets("PE Review").Select
         popwb.Sheets("PE Review").Copy before:=wbfile.Sheets(1)
    End Select
    
    Application.DisplayAlerts = False
  
  'deleting extra sheets
    wbfile.Sheets("Sheet1").Delete
    wbfile.Sheets("Sheet2").Delete
    wbfile.Sheets("Sheet3").Delete
    
    Application.DisplayAlerts = True
  
  'formatting month
    monthstr = Application.WorksheetFunction.Text(revmonth, "YYYYMM")
    
    'oldStatusBar = Application.DisplayStatusBar
    Application.DisplayStatusBar = True

 'adding sheets depending on how many review schedules there are for each program.
 'Also adding in information from temp file
   Select Case program
    Case "GA"
        Call Ga
         
    Case "MA Positive"
        Call Mapos
         
    Case "FS Positive", "FS Supplemental"
        Call Fspos
         
    Case "FS Negative"
        Call Fsneg
         
    Case "TANF", "TANF CAR"
        Call tanf
         
    Case "MA Closed", "MA CHIP Closed", "MA Rejected", "MA CHIP Rejected"
        Call Maneg
        
    Case "MA PE"
        Call PE

    End Select
    
   
   'Dim sh As Worksheet
   'Dim wb As Workbook
   'Dim OutApp As Object
   'Dim OutMail As Object
   Dim rngSource As Range, rngTarget As Range
   Set Destwb = Workbooks(fileName)
   Workbooks(fileName).SAVE
   
   'Set OutApp = CreateObject("Outlook.Application")
   'OutApp.Session.Logon
   'Set OutMail = OutApp.CreateItem(0)

 'cleaning out columns that have file names in them
   Windows("populate.xlsm").Activate
   Sheets("Populate").Select
   Range("AL2:BY40").Clear
   
   shtotalnumber = Workbooks(fileName).Worksheets.Count - 1
   shnumber = 0
   
 For Each Sh In Workbooks(fileName).Worksheets
     If Sh.Name <> "Temp" Then
     shnumber = shnumber + 1

    strTemp = "Creating Schedule " & Sh.Name & " - " & shnumber & "/" & shtotalnumber & ". Please be patient..."
    Application.StatusBar = strTemp

  Select Case program
   Case "GA", "TANF", "TANF CAR", "MA Positive"
       examnum = Sh.Range("AO3") & Sh.Range("AP3")
    Case "FS Positive", "FS Supplemental"
        examnum = Sh.Range("AJ5") & Sh.Range("AK5")
    Case "FS Negative"
        examnum = Sh.Range("W17")
    Case "MA Closed", "MA CHIP Closed", "MA Rejected", "MA CHIP Rejected"
        examnum = Sh.Range("AB11") & Sh.Range("AC11")
    Case "MA PE"
        examnum = Sh.Range("AB11") & Sh.Range("AC11")
 End Select
        
          Sh.Copy
          Set wb = ActiveWorkbook
 'create workbook containing schedule
           tempfilename = "Review Number " & Sh.Name & " Month " _
                         & monthstr & " Examiner " & examnum & ".xlsm"
           Application.DisplayAlerts = False
              wb.SaveAs fileName:=tempfilename, FileFormat:=52
            Windows("populate.xlsm").Activate
            
Select Case program
    Case "GA"
         Sheets("GA Computation").Select
         Sheets("GA Computation").Copy before:=Workbooks(tempfilename).Sheets(1)
         Sh.Range("B2").Copy
         Sheets("GA Computation").Range("B1").Select
         Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
         Sh.Range("A10").Copy
         Sheets("GA Computation").Range("e2").Select
         Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
         tempstr1 = Right(Sh.Range("U10"), 2)
         Sheets("GA Computation").Range("C2") = tempstr1
        'Windows("populate.xlsm").Activate
         Workbooks.Open fileName:=strcurdir & "GA Workbook.XLSM", UpdateLinks:=False
         Sheets("GA Workbook").Select
         Sheets("GA Workbook").Copy before:=Workbooks(tempfilename).Sheets(1)
         Workbooks("GA Workbook.XLSM").Close False
         Sheets("GA Workbook").Range("b13") = Sh.Range("O4")
         Sheets("GA Workbook").Range("b15") = Sh.Range("B2")
         Sheets("GA Workbook").Range("b17") = Sh.Range("B3")
         Sheets("GA Workbook").Range("b18") = Sh.Range("B4")
         Sheets("GA Workbook").Range("b19") = Sh.Range("B5")
         Sheets("GA Workbook").Range("g21") = Sh.Range("I10")
         Sheets("GA Workbook").Range("g22") = Sh.Range("A10")
         Sheets("GA Workbook").Range("AI1") = Sh.Range("A10")
         Sheets("GA Workbook").Range("g23") = Sh.Range("AB10")
         Sheets("GA Workbook").Range("g32") = Sh.Range("AO3") & Sh.Range("AP3")
         'finding maxrow in temp
         For i = 1 To maxrow
         'Bringing over line number
            jrow = 11
            jcol = 30
            'If Workbooks(Filename).Sheets("Temp").Cells(i, jcol) = "" Then jrow = ""
            nameflag = 0
            pmtname = Workbooks(fileName).Sheets("Temp").Cells(i, 9) & " " & Workbooks(fileName).Sheets("Temp").Cells(i, 8)
            pmtname2 = wb.Sheets(Sh.Name).Range("b2")
    If pmtname = pmtname2 Then
            Do Until (Workbooks(fileName).Sheets("Temp").Cells(i, jcol) = "" Or jrow > 22)
                If Workbooks(fileName).Sheets("Temp").Cells(i, jcol) < 10 Then
                    Sheets("GA Workbook").Cells(jrow, 10) = "0" & Workbooks(fileName).Sheets("Temp").Cells(i, jcol)
                Else
                    Sheets("GA Workbook").Cells(jrow, 10) = Workbooks(fileName).Sheets("Temp").Cells(i, jcol)
                End If
         'Bringing First Name and Last Name
          'Sheets("TANF GA Face Sheet").Cells(jrow, 16) = pmtname2
           Sheets("GA Workbook").Cells(jrow, 12) = Workbooks(fileName).Sheets("Temp").Cells(i, jcol + 3) & " " & Workbooks(fileName).Sheets("Temp").Cells(i, jcol + 2)
          If pmtname = Sheets("GA Workbook").Cells(jrow, 12) Then
            nameflag = 1
          End If
         'Bringing over Birthdate
          Sheets("GA Workbook").Cells(jrow, 22) = Workbooks(fileName).Sheets("Temp").Cells(i, jcol + 4)
         'Bringing over Social Security Number
          Sheets("GA Workbook").Cells(jrow, 31) = Workbooks(fileName).Sheets("Temp").Cells(i, jcol + 5)
         'Bringing over Cat.
          Sheets("GA Workbook").Cells(jrow, 29) = Workbooks(fileName).Sheets("Temp").Cells(i, 7)
         'Putting Y in Recip
          Sheets("GA Workbook").Cells(jrow, 36) = "Y"
         'Calculating age - use last day of month before review month of GA
                prevmonth = CInt(Left(Sh.Range("ab10"), 2)) - 1
                reviewyear = CInt(Right(Sh.Range("ab10"), 4))
                If prevmonth = 0 Then
                    prevmonth = 12
                    reviewyear = reviewyear - 1
                End If
                Sheets("GA Workbook").Cells(jrow, 25) = Int((EOMONTH(DateSerial(reviewyear, prevmonth, 1)) - Workbooks(fileName).Sheets("Temp").Cells(i, jcol + 4)) / 365)
                jrow = jrow + 1
                jcol = jcol + 6
            Loop
         'Figuring out if Payment Name is already in excel spreadsheet
         If nameflag = 0 Then
            Range("l11:ak21").Copy
            Range("l12:ak22").PasteSpecial (xlPasteAll)
            Range("l11:ak11").ClearContents
            Range("l11") = pmtname
         End If
         Exit For
    End If
            Next i
            
    'MA Positive
            
    Case "MA Positive"
    
        'Insert and Populate consent form
         Workbooks.Open fileName:=strcurdir & "MA Consent.XLSX", UpdateLinks:=False
         
         Sheets("Consent").Select
         Sheets("Consent").Copy After:=Workbooks(tempfilename).Sheets(1)
         Workbooks("MA Consent.XLSX").Close False
         'Name
         Sheets("Consent").Range("A8") = Sh.Range("B2")
         'County
         Sheets("Consent").Range("H4") = Right(Sh.Range("AC10"), 2)
         'Record Number
         Sheets("Consent").Range("I4") = Sh.Range("H10")
         'Category
         tempcat = Sh.Range("Q10") & "-" & Sh.Range("V10")
         If Sh.Range("Z10") <> "" Then
            tempcat = tempcat & "-" & Sh.Range("Z10")
        End If
         Sheets("Consent").Range("J4") = tempcat
         'Review Number
         Sheets("Consent").Range("H6") = Sh.Range("A10") & Sh.Range("B10") & Sh.Range("c10") & Sh.Range("D10") & Sh.Range("E10") & Sh.Range("F10")
         'District
         Sheets("Consent").Range("L4") = Sh.Range("AH10")
    
         'Insert survey sheet
    '     Windows("populate.xlsm").Activate
    '     If Sh.Range("O4") = "Philadelphia County , Snyder District" Or Sh.Range("O4") = "Cumberland County" Or Sh.Range("O4") = "Dauphin County" Then
    '     Sheets("SurveyCopy").Select
    '    Else
    '     Sheets("Survey").Select
    '    End If
    '    ActiveSheet.Copy After:=Workbooks(TempFileName).Sheets(1)
    '     ActiveSheet.Name = "Survey"
    '     Sh.Range("AF2").Copy
    '     Sheets("Survey").Range("C24").PasteSpecial
    '     Sheets("Survey").Range("C35").PasteSpecial
    '     Sheets("Survey").Range("C46").PasteSpecial
         'sh.Range("A18").Copy
         'Sheets("Survey").Range("K174").Select
         'Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            ':=False, Transpose:=False
         'ActiveSheet.Copy After:=Workbooks(TempFileName).Sheets(1)
    '     TempStr = Sh.Range("A10") & Sh.Range("B10") & Sh.Range("C10") & Sh.Range("D10") & Sh.Range("E10") & Sh.Range("F10")
    '     Sheets("Survey").Range("K175") = TempStr
         
         'copy over 721 and 3rd party worksheets
         'Windows("populate.xlsm").Activate
         'Sheets("PA721").Select
         'Sheets("PA721").Copy After:=Workbooks(tempfilename).Sheets(1)
        ' Windows("populate.xlsm").Activate
        ' Sheets("3rd party").Select
        ' Sheets("3rd party").Copy After:=Workbooks(tempfilename).Sheets(1)
         
         'Bringing over name
         'Sheets("PA721").Range("c2") = Sh.Range("b2")
        ' Sheets("3rd party").Range("a7") = Sh.Range("b2")
        ' Sheets("3rd party").Range("E61") = Sh.Range("b2")
         'Bringing over case number
         'Sheets("PA721").Range("ag2") = Right(Sh.Range("AC10"), 2) & "/" & Sh.Range("h10")
         'Bringing over Category and Program Status Code and Grant Group if not blank
         'If Sh.Range("Z10") = "" Then
         '   Sheets("PA721").Range("AG3") = Sh.Range("Q10") & "-" & Sh.Range("V10")
         '   Sheets("3rd party").Range("G7") = Sh.Range("h10") & "  " & Sh.Range("Q10") & "-" & Sh.Range("V10")
         'Else
         '   Sheets("PA721").Range("AG3") = Sh.Range("Q10") & "-" & Sh.Range("V10") & "-" & Sh.Range("Z10")
         '   Sheets("3rd party").Range("G7") = Sh.Range("h10") & "  " & Sh.Range("Q10") & "-" & Sh.Range("V10") & "-" & Sh.Range("Z10")
         'End If
         'copying Category etc to another cell
         'Sheets("PA721").Range("A33").value = Sheets("PA721").Range("AG3").value
        ' Sheets("3rd party").Range("H61") = Sheets("3rd party").Range("G7")
         'Bringing over review number
         'Sheets("PA721").Range("al2") = Sh.Range("a10") & Sh.Range("b10") & Sh.Range("c10") & Sh.Range("d10") & Sh.Range("e10") & Sh.Range("f10")
        ' Sheets("3rd party").Range("G4") = Sh.Range("a10") & Sh.Range("b10") & Sh.Range("c10") & Sh.Range("d10") & Sh.Range("e10") & Sh.Range("f10")
        ' Sheets("3rd party").Range("H59") = Sheets("3rd party").Range("G4")
         'Bringing over review month
         'Sheets("PA721").Range("al4") = Sh.Range("am10")
        ' Sheets("3rd party").Range("D11") = Sh.Range("am10")
         'Bringing over reviewer name
         'Sheets("PA721").Range("al6") = Sh.Range("af2") & " - " & Sh.Range("ao3") & Sh.Range("ap3")
        ' Sheets("3rd party").Range("d4") = Sh.Range("af2") & " - " & Sh.Range("ao3") & Sh.Range("ap3")
         'Bringing over first line in address
         'Sheets("PA721").Range("a4") = Sh.Range("b3")
        ' Sheets("3rd party").Range("a8") = Sh.Range("b3")
         'Bringing over second line in address
         'Sheets("PA721").Range("a5") = Sh.Range("b4")
        ' If Sh.Range("b4") = "" Then
        '    Sheets("3rd party").Range("a9") = Sh.Range("b5")
        ' Else
        '    Sheets("3rd party").Range("a9") = Sh.Range("b4") & ", " & Sh.Range("b5")
        ' End If
         'Bringing over third line in address
         'Sheets("PA721").Range("a6") = Sh.Range("b5")
         'Individual Reviewing
         'Sheets("PA721").Range("D19") = Workbooks(FileName).Sheets("Temp").Cells(i, 27)
         'Claim Amount
         'Sheets("PA721").Range("O19") = sh.Range("AN16")
        'Putting Certification Date
         'revmn = DateSerial(Val(Right(Sheets("PA721").Range("AL4"), 4)), Val(Left(Sheets("PA721").Range("AL4"), 2)), 1) 'first day of sample month
         'Sheets("PA721").Range("P3") = revmn 'first day of sample month
         'Sheets("PA721").Range("P3") = Evaluate("=EOMONTH(revmn, -1)+1") 'first day of sample month
         'Sheets("PA721").Range("P4") = EOMONTH(revmn, 0) 'last day of sample month
         'Sheets("PA721").Range("P3").NumberFormat = "m/dd/yyyy"
         'Sheets("PA721").Range("P4").NumberFormat = "m/dd/yyyy"
         'Sheets("PA721").Range("Y3").NumberFormat = "m/dd/yyyy"
         'Sheets("PA721").Range("Y4").NumberFormat = "m/dd/yyyy"

         'Putting Check marks for certification periods
         'Cat = Sheets("PA721").Range("AG3").value
         'If Cat = "PAN-00" Or Cat = "PAN-80" Or Cat = "PJN-00" Or Cat = "PJN-80" Or Cat = "PAN-66" Or Cat = "PJN-66" Then
         '    Sheets("PA721").Shapes("Check Box 8").OLEFormat.Object.value = 1
         'Else
         '   Sheets("PA721").Shapes("Check Box 9").OLEFormat.Object.value = 1
         'End If
         'Bringing over State and County Code
        ' Sheets("3rd party").Range("b11") = "42" & "/" & Right(Sh.Range("AC10"), 2)
         
        'paste in MA resource computation sheet
         Workbooks("Populate.xlsm").Sheets("MA Resources").Copy before:=Workbooks(tempfilename).Sheets(1)
         'Sheets("MA Resources").Copy before:=Workbooks(TempFileName).Sheets(1)
         Sh.Range("B2").Copy
         Sheets("MA Resources").Range("B1").Select
         Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
         Sh.Range("AM10").Copy
         Sheets("MA Resources").Range("I1").Select
         Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
         tempstr1 = Sh.Range("A10") & Sh.Range("B10") & Sh.Range("C10") & Sh.Range("D10") & Sh.Range("E10") & Sh.Range("F10")
         Sheets("MA Resources").Range("F1") = tempstr1
         
         'Insert MA Income worksheet
         Workbooks.Open fileName:=strcurdir & "MA INCOME COMPUTATION FORM.xlsm", UpdateLinks:=False
         
         Sheets("MA Income Comp").Select
         Sheets("MA Income Comp").Copy After:=Workbooks(tempfilename).Sheets(1)
         Workbooks("MA INCOME COMPUTATION FORM.xlsm").Close False
         'Name
         Sheets("MA Income Comp").Range("A4") = Sh.Range("B2")
         Sheets("MA Income Comp").Range("N4") = Sh.Range("B2")
         'Review Number
         Sheets("MA Income Comp").Range("J5") = Sh.Range("A10") & Sh.Range("B10") & Sh.Range("C10") & Sh.Range("D10") & Sh.Range("E10") & Sh.Range("F10")
         Sheets("MA Income Comp").Range("W5") = Sh.Range("A10") & Sh.Range("B10") & Sh.Range("C10") & Sh.Range("D10") & Sh.Range("E10") & Sh.Range("F10")
         'Bringing over review month
         Sheets("MA Income Comp").Range("J7") = Sh.Range("AM10")
         Sheets("MA Income Comp").Range("W7") = Sh.Range("AM10")
         Sheets("MA Income Comp").Range("C13") = Sh.Range("AM10")
         Sheets("MA Income Comp").Range("D13") = Sh.Range("AM10")
         Sheets("MA Income Comp").Range("P15") = Sh.Range("AM10")
         Sheets("MA Income Comp").Range("Q15") = Sh.Range("AM10")
         'Bringing over Cat. and Program Status Code and Grant Group if not blank
         If Sh.Range("Z10") = "" Then
            tempstrcpg = Sh.Range("Q10") & "-" & Sh.Range("V10")
         Else
            tempstrcpg = Sh.Range("Q10") & "-" & Sh.Range("V10") & "-" & Sh.Range("Z10")
         End If
         Sheets("MA Income Comp").Range("C12") = tempstrcpg
         Sheets("MA Income Comp").Range("D12") = tempstrcpg
         Sheets("MA Income Comp").Range("P14") = tempstrcpg
         Sheets("MA Income Comp").Range("Q14") = tempstrcpg
         'Bringing over Certification Period
        
         'copy over ma workbook
         Workbooks.Open fileName:=strcurdir & "MA Workbook.XLSM", UpdateLinks:=False
         Sheets("MA Workbook").Select
         Sheets("MA Workbook").Copy before:=Workbooks(tempfilename).Sheets(1)
         Workbooks("MA Workbook.XLSM").Close False
         'Windows("populate.xlsm").Activate
         'Sheets("MA Face Sheet").Select
         'Sheets("MA Face Sheet").Copy After:=Workbooks(TempFileName).Sheets(1)
         'wb.sh.Range("a1").Select
         'Bringing things over from Schedule to MA Workbook
    
         Sheets("MA Workbook").Range("E12") = Right(Sh.Range("AC10"), 2)
         Sheets("MA Workbook").Range("G12") = Sh.Range("AH10")
         
         Sheets("MA Workbook").Range("C13") = Sh.Range("O4")
         Sheets("MA Workbook").Range("C15") = Sh.Range("B2")
         Sheets("MA Workbook").Range("C17") = Sh.Range("B3")
         Sheets("MA Workbook").Range("C18") = Sh.Range("B4")
         Sheets("MA Workbook").Range("C19") = Sh.Range("B5")
         'Review Number
         Sheets("MA Workbook").Range("F23") = Sh.Range("A10") & Sh.Range("B10") & Sh.Range("C10") & Sh.Range("D10") & Sh.Range("E10") & Sh.Range("F10")
        

         'Bringing over multiple items
         If Sh.Range("Z10") <> "" Then
            mastr = Sh.Range("Q10") & "-" & Sh.Range("V10") & "-" & Sh.Range("Z10")
         Else
            mastr = Sh.Range("Q10") & "-" & Sh.Range("V10")
         End If
         Sheets("MA Workbook").Range("F21") = Sh.Range("H10") ' Case Number
         Sheets("MA Workbook").Range("F22") = mastr 'Category, Program Status Code, GG
         Sheets("MA Workbook").Range("F24") = Sh.Range("AM10") ' Review Month
         Sheets("MA Workbook").Range("F33") = Sh.Range("AF2") & " - " & Sh.Range("AO3") & Sh.Range("AP3") 'Reviewer Name & Number
         'Bringing over Certification Period on income sheet
         'Sheets("MA Workbook").Range("G31").Copy
         'Sheets("MA Income Comp").Select
         'Sheets("MA Income Comp").Range("H4").Select
         'Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
         '   :=False, Transpose:=False
         'Sheets("MA Income Comp").Range("H6").Select
         'Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
         '   :=False, Transpose:=False
         'Sheets("MA Workbook").Range("G32").Copy
         'Sheets("MA Income Comp").Range("H5").Select
         'Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
         '   :=False, Transpose:=False
         'Sheets("MA Income Comp").Range("H7").Select
         'Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
          '  :=False, Transpose:=False
          'Sheets("MA Income Comp").Range("H4") = Sheets("MA Workbook").Range("G31")
          'Sheets("MA Income Comp").Range("H5") = Sheets("MA Workbook").Range("G32")
          Sheets("MA Income Comp").Range("H6") = Sheets("MA Workbook").Range("G31")
          Sheets("MA Income Comp").Range("H7") = Sheets("MA Workbook").Range("G32")
          Sheets("MA Income Comp").Range("T10") = Sheets("MA Workbook").Range("AC37")
          
         'Sheets("MA Workbook").Range("AH1") = Sheets("MA Workbook").Range("F22")
         'finding maxrow in temp
         For i = 1 To maxrow
         'Bringing over line number
         'jrow is coming from MA facesheet
            jrow = 11
         'jcol = line number on temp
            'jcol = 30
            jcol = 32
            'If Workbooks(Filename).Sheets("Temp").Cells(i, jcol) = "" Then jrow = ""
            nameflag = 0
            'pmtname = Workbooks(FileName).Sheets("Temp").Cells(i, 9) & " " & Workbooks(FileName).Sheets("Temp").Cells(i, 8)
            pmtname = wb.Sheets(Sh.Name).Range("b2")
            pmtname2 = Workbooks(fileName).Sheets("Temp").Cells(i, jcol)
            'pmtname2 = wb.Sheets(Sh.Name).Range("b2")
    If pmtname = pmtname2 Then
            Do Until (Workbooks(fileName).Sheets("Temp").Cells(i, jcol) = "" Or jrow > 22)
            'For Line Number
               ' If Workbooks(FileName).Sheets("Temp").Cells(i, jcol) < 10 Then
               '     Sheets("MA Workbook").Cells(jrow, 10) = "0" & Workbooks(FileName).Sheets("Temp").Cells(i, jcol)
               ' Else
               '     Sheets("MA Workbook").Cells(jrow, 10) = Workbooks(FileName).Sheets("Temp").Cells(i, jcol)
               ' End If
         'Bringing First Name and Last Name
          Sheets("MA Workbook").Cells(jrow, 12) = pmtname
          'Sheets("MA Workbook").Cells(jrow, 12) = Workbooks(FileName).Sheets("Temp").Cells(i, jcol + 3) & " " & Workbooks(FileName).Sheets("Temp").Cells(i, jcol + 2)
         
         'Figuring out if Payment Name is already in excel spreadsheet
          
         ' If pmtname2 = Sheets("MA Workbook").Cells(jrow, 12)
          'If pmtname2 = Sheets("MA Workbook").Cells(jrow, 12) Then
            nameflag = 1
         ' End If
         'Bringing over Birthdate
          'Sheets("MA Workbook").Cells(jrow, 22) = Workbooks(FileName).Sheets("Temp").Cells(i, jcol + 4)
          Sheets("MA Workbook").Cells(jrow, 22) = Workbooks(fileName).Sheets("Temp").Cells(i, 34)
         'Bringing over Social Security Number
          'Sheets("MA Workbook").Cells(jrow, 31) = Workbooks(FileName).Sheets("Temp").Cells(i, jcol + 5)
          Sheets("MA Workbook").Cells(jrow, 31) = Workbooks(fileName).Sheets("Temp").Cells(i, 35)
         'Bringing over Cat.
          Sheets("MA Workbook").Cells(jrow, 24) = mastr
         'Putting Y in Recip
          Sheets("MA Workbook").Cells(jrow, 36) = "Y"
         'Calculating age - use last day of review month for MA
'                Sheets("MA Workbook").Cells(jrow, 25) = Int((EOMONTH(DateSerial(Right(Sh.Range("am10"), 4), Left(Sh.Range("am10"), 2), 1)) - Workbooks(FileName).Sheets("Temp").Cells(i, jcol + 4)) / 365)
                jrow = jrow + 1
                jcol = jcol + 7
            Loop
         'Figuring out if Payment Name is already in excel spreadsheet
         Sheets("MA Workbook").Select
         If nameflag = 0 Then
            Range("J11:AJ21").Copy
            Range("J12:AJ22").PasteSpecial (xlValue)
            Range("J11:AJ11").ClearContents
            Range("L11") = pmtname
            Range("J11").Select
         End If
            
        If Sheets("MA Workbook").Range("AM4") = "" Then
            Sheets("MA Workbook").Range("AM4") = Format(Date, "short date")
        End If
            
        'put this name in the 721 sheet
         'If Sheets("MA Workbook").Range("L12") = "" Then
         '   Sheets("PA721").Range("A19") = Sheets("MA Workbook").Range("J11")
         '   Sheets("PA721").Range("D19") = Sheets("MA Workbook").Range("L11")
         'Else
         '   Sheets("PA721").Range("A19") = Sheets("MA Workbook").Range("J12")
         '   Sheets("PA721").Range("D19") = Sheets("MA Workbook").Range("L12")
         'End If

          Exit For
    End If
        Sheets(Sh.Name).Select
        Range("A1").Select
    Next i
   
        'Populate Tracking form
         Workbooks.Open fileName:=strcurdir & "MA Tracking.XLSX", UpdateLinks:=False
         
         Sheets("MA Tracking").Select
         Sheets("MA Tracking").Copy before:=Workbooks(tempfilename).Sheets(1)
         Workbooks("MA Tracking.XLSX").Close False
         'Name
         Sheets("MA Tracking").Range("B6") = Sh.Range("B2")
         'County
         Sheets("MA Tracking").Range("B4") = Sh.Range("O4")
         'Record Number-Category-PS
         Sheets("MA Tracking").Range("J6") = Sh.Range("H10") & "-" & Sh.Range("Q10") & "-" & Sh.Range("V10")
         'Examiner Number
         Sheets("MA Tracking").Range("M4") = Sh.Range("AO3") & Sh.Range("AP3")
         'Review Number
         Sheets("MA Tracking").Range("J4") = Sh.Range("A10") & Sh.Range("B10") & Sh.Range("c10") & Sh.Range("D10") & Sh.Range("E10") & Sh.Range("F10")
         'Sample Month
         Sheets("MA Tracking").Range("G4") = Sh.Range("AM10")
         'Address
         Sheets("MA Tracking").Range("B8") = Sh.Range("B3")
         Sheets("MA Tracking").Range("B9") = Sh.Range("B4")
         Sheets("MA Tracking").Range("B10") = Sh.Range("B5")
         
    'FS Positive and FS Supplemental
    Case "FS Positive", "FS Supplemental"
      '  If Sh.Range("M5") = "Philadelphia County , Snyder District" Or Sh.Range("M5") = "Cumberland County" Or Sh.Range("M5") = "Dauphin County" Then
      '   Sheets("SurveyCopy").Select
      '  Else
      '   Sheets("Survey").Select
      '  End If
      '   ActiveSheet.Copy After:=Workbooks(TempFileName).Sheets(1)
      '   ActiveSheet.Name = "Survey"
      '   Sh.Range("AE5").Copy
      '   Sheets("Survey").Range("C24").PasteSpecial
      '   Sheets("Survey").Range("C35").PasteSpecial
      '   Sheets("Survey").Range("C46").PasteSpecial
      '   Sh.Range("A18").Copy
      '   Sheets("Survey").Range("K174").Select
      '   Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      '      :=False, Transpose:=False
         Windows("populate.xlsm").Activate
         Sheets("FS Computation").Select
         Sheets("FS Computation").Copy before:=Workbooks(tempfilename).Sheets(1)
         Sh.Range("B4").Copy
         Sheets("FS Computation").Range("B1").Select
         Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
         Sh.Range("A18").Copy
         Sheets("FS Computation").Range("F1").Select
         Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
         tempstr1 = Sh.Range("AD18") & "/" & Sh.Range("AG18")
         Sheets("FS Computation").Range("H1") = tempstr1
      'Populate consent form
         Workbooks.Open fileName:=strcurdir & "Consent.XLSX", UpdateLinks:=False
         
         Sheets("Consent").Select
         Sheets("Consent").Copy before:=Workbooks(tempfilename).Sheets(1)
         'Name
         Sheets("Consent").Range("A8") = Sh.Range("B4")
         'County
         Sheets("Consent").Range("H4") = Right(Sh.Range("X18"), 2)
         'Record Number
         Sheets("Consent").Range("I4") = Left(Sh.Range("I18"), 9)
         'Category
         Sheets("Consent").Range("J4") = "FS"
         'Review Number
         Sheets("Consent").Range("H6") = Sh.Range("A18")
         'District
         Sheets("Consent").Range("L4") = Sh.Range("B153") & Sh.Range("C153")
         'First Address
         Sheets("Consent").Range("A9") = Sh.Range("B5")
         'Second Address
         Sheets("Consent").Range("A10") = Sh.Range("B6")
         'City, State, Zip
         Sheets("Consent").Range("A11") = Sh.Range("B7")
         Workbooks("Consent.XLSX").Close False
         
        'Open FS workbook
        Workbooks.Open fileName:=strcurdir & "FS Workbook.XLSM", UpdateLinks:=False
         
         Sheets("FS Workbook").Select
         Sheets("FS Workbook").Copy before:=Workbooks(tempfilename).Sheets(1)
         'name
         Sheets("FS Workbook").Range("b11") = Sh.Range("b4")
         'county and district
         Sheets("FS Workbook").Range("b9") = Sh.Range("m5")
         'address
         Sheets("FS Workbook").Range("b13") = Sh.Range("B5")
         Sheets("FS Workbook").Range("b14") = Sh.Range("b6")
         Sheets("FS Workbook").Range("b15") = Sh.Range("b7")
         'review Number
         Sheets("FS Workbook").Range("g18") = Sh.Range("A18")
         Sheets("FS Workbook").Range("AI1") = Sh.Range("A18")
         'case number
         Sheets("FS Workbook").Range("G17") = Sh.Range("I18")
         'reviewer
         Sheets("FS Workbook").Range("g32") = Sh.Range("AE5") & " - " & Sh.Range("AJ5") & Sh.Range("AK5")
         'review month
         Sheets("FS Workbook").Range("g19") = Sh.Range("AD18") & "/" & Sh.Range("AG18")
         'reviewer name
         tempArray = Split(Sh.Range("AE5").Value, " ")
                            JustLastName = tempArray(UBound(tempArray))
         Sheets("FS Workbook").Range("g32") = JustLastName & "-" & Sh.Range("AJ5") & Sh.Range("AK5")
         Workbooks("FS Workbook.XLSM").Close False
    
        If Sheets("FS Workbook").Range("AL4") = "" Then
            Sheets("FS Workbook").Range("AL4") = Format(Date, "short date")
        End If
    
    Case "TANF CAR"
         Sheets("TANF CAR").Select
         Sheets("TANF CAR").Copy before:=Workbooks(fileName).Sheets(1)
         
    Case "FS Negative"
         Windows("populate.xlsm").Activate
         Sheets("FS Computation").Select
         Sheets("FS Computation").Copy before:=Workbooks(tempfilename).Sheets(1)
         'Name
         Sh.Range("C8").Copy
         Sheets("FS Computation").Range("B1").Select
         Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
         'Review Number
         Sh.Range("C20").Copy
         Sheets("FS Computation").Range("F1").Select
         Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
         'Sample Month
         tempstr1 = Sh.Range("AF20") & "/" & Sh.Range("AI20")
         Sheets("FS Computation").Range("H1") = tempstr1
    
    Case "TANF"
        'copy comp sheet
         Sheets("TANF Computation").Select
         Sheets("TANF Computation").Copy before:=Workbooks(tempfilename).Sheets(1)
         Sh.Range("B2").Copy 'client name
         Sheets("TANF Computation").Range("B1").Select
         Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
         Sh.Range("A10").Copy 'review number
         Sheets("TANF Computation").Range("e2").Select
         Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
         tempstr1 = Right(Sh.Range("U10"), 2) 'county number
         Sheets("TANF Computation").Range("C2") = tempstr1
         'Windows("populate.xlsm").Activate
         
         'copy workbook
         Workbooks.Open fileName:=strcurdir & "TANF Workbook.XLSM", UpdateLinks:=False
         Sheets("TANF Workbook").Select
         Sheets("TANF Workbook").Copy before:=Workbooks(tempfilename).Sheets(1)
         CopyModule Workbooks("TANF Workbook.XLSM"), "TANF_WS", Workbooks(tempfilename)

         Workbooks("TANF Workbook.XLSM").Close False
         'Sheets("TANF Workbook").CheckBoxes(1).Value = xlOn
         Sheets("TANF Workbook").Range("b13") = Sh.Range("O4") 'county/district
         Sheets("TANF Workbook").Range("b15") = Sh.Range("B2") 'client name
         Sheets("TANF Workbook").Range("b17") = Sh.Range("B3") 'client 1st address
         Sheets("TANF Workbook").Range("b19") = Sh.Range("B4") 'client 2nd address
         Sheets("TANF Workbook").Range("g21") = Sh.Range("I10") 'case number
         Sheets("TANF Workbook").Range("g22") = Sh.Range("A10") 'review number
         Sheets("TANF Workbook").Range("AH1") = Sh.Range("A10") 'review number
         Sheets("TANF Workbook").Range("g23") = Sh.Range("AB10") 'sample month
         Sheets("TANF Workbook").Range("g32") = Sh.Range("AO3") & Sh.Range("AP3") 'examiner number
         'finding maxrow in temp
         For i = 1 To maxrow
         'Bringing over line number
            jrow = 11
            jcol = 30
            'If Workbooks(Filename).Sheets("Temp").Cells(i, jcol) = "" Then jrow = ""
            nameflag = 0
            pmtname = Workbooks(fileName).Sheets("Temp").Cells(i, 9) & " " & Workbooks(fileName).Sheets("Temp").Cells(i, 8)
            pmtname2 = wb.Sheets(Sh.Name).Range("b2")
    If pmtname = pmtname2 Then
            Do Until (Workbooks(fileName).Sheets("Temp").Cells(i, jcol) = "" Or jrow > 22)
                If Workbooks(fileName).Sheets("Temp").Cells(i, jcol) < 10 Then
                    Sheets("TANF Workbook").Cells(jrow, 10) = "0" & Workbooks(fileName).Sheets("Temp").Cells(i, jcol)
                Else
                    Sheets("TANF Workbook").Cells(jrow, 10) = Workbooks(fileName).Sheets("Temp").Cells(i, jcol)
                End If
         'Bringing First Name and Last Name
          Sheets("TANF Workbook").Cells(jrow, 12) = Workbooks(fileName).Sheets("Temp").Cells(i, jcol + 3) & " " & Workbooks(fileName).Sheets("Temp").Cells(i, jcol + 2)
          
         'Figuring out if Payment Name is already in excel spreadsheet
          If pmtname = Sheets("TANF Workbook").Cells(jrow, 12) Then
            nameflag = 1
          End If
         'Bringing over Birthdate
          Sheets("TANF Workbook").Cells(jrow, 22) = Workbooks(fileName).Sheets("Temp").Cells(i, jcol + 4)
         'Bringing over Social Security Number
          Sheets("TANF Workbook").Cells(jrow, 31) = Workbooks(fileName).Sheets("Temp").Cells(i, jcol + 5)
         'Bringing over Cat.
          Sheets("TANF Workbook").Cells(jrow, 29) = Workbooks(fileName).Sheets("Temp").Cells(i, 7)
         'Putting Y in Recip
          Sheets("TANF Workbook").Cells(jrow, 35) = "Yes"
         'Calculating age - use last day of month before review month of TANF
                prevmonth = CInt(Left(Sh.Range("ab10"), 2)) - 1
                reviewyear = CInt(Right(Sh.Range("ab10"), 4))
                If prevmonth = 0 Then
                    prevmonth = 12
                    reviewyear = reviewyear - 1
                End If
                Age = Int((EOMONTH(DateSerial(reviewyear, prevmonth, 1)) - Workbooks(fileName).Sheets("Temp").Cells(i, jcol + 4)) / 365)
                If Age < 10 Then
                    textage = "0" & Age
                Else
                    textage = Age
                End If
                Sheets("TANF Workbook").Cells(jrow, 25) = textage
                jrow = jrow + 1
                jcol = jcol + 6
            Loop
         'Figuring out if Payment Name is already in excel spreadsheet
         If nameflag = 0 Then
            Range("j11:aj21").Copy
            Range("j12:aj22").PasteSpecial (xlPasteAll)
            Range("j11:aj11").ClearContents
            Range("l11") = pmtname
         End If
         Exit For
    End If
            Next i
            
        If Sheets("TANF Workbook").Range("AI3") = "" Then
            Sheets("TANF Workbook").Range("AI3") = Format(Date, "short date")
        End If
            
        'make merge cells in section B of workbook
        Range("J11:AH11").Copy
            
        For i = 12 To 22
            Range("J" & i & ":AH" & i).Select
            Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                SkipBlanks:=False, Transpose:=False
        Next i
        Application.CutCopyMode = False
    
        'copy AMR worksheet
        'popwb.Sheets("AMR").Copy after:=Workbooks(TempFileName).Sheets("TANF Coumputation")
        'Workbooks(TempFileName).Sheets("AMR").Range("C2") = Sh.Range("A10") 'county/district
        'Workbooks(TempFileName).Sheets("AMR").Range("C3") = Sh.Range("O4") 'county/district
        'Workbooks(TempFileName).Sheets("AMR").Range("C4") = Sh.Range("AB10") 'sample month
        'Workbooks(TempFileName).Sheets("AMR").Range("C5") = Sh.Range("I10") 'case number
        'Workbooks(TempFileName).Sheets("AMR").Range("C7") = Sh.Range("B2") 'client name
        'Workbooks(TempFileName).Sheets("AMR").Range("C8") = Sh.Range("AO3") & Sh.Range("AP3") 'examiner number
        
    End Select
    
Dim sourceBook As Workbook
Dim destinationBook As Workbook
Dim codeFromSource As String
 
'//Tested early and late-bound (removed 5.3)//
Dim DestCom As Object
Dim DestMod As Object
 
    'copy macros to new workbook and assign macros to buttons in new workbook
    CopyModule Workbooks("populate.xlsm"), "Module1", Workbooks(tempfilename)
    CopyModule Workbooks("populate.xlsm"), "TANFMod", Workbooks(tempfilename)
    CopyModule Workbooks("populate.xlsm"), "Module3", Workbooks(tempfilename)
    CopyModule Workbooks("populate.xlsm"), "Finding_Memo", Workbooks(tempfilename)
    CopyModule Workbooks("populate.xlsm"), "CAO_Appointment", Workbooks(tempfilename)
    CopyForm Workbooks("populate.xlsm"), "SelectDate", Workbooks(tempfilename)
    CopyForm Workbooks("populate.xlsm"), "SelectTime", Workbooks(tempfilename)
    CopyForm Workbooks("populate.xlsm"), "SelectForms", Workbooks(tempfilename)
    CopyModule Workbooks("populate.xlsm"), "CashMemos", Workbooks(tempfilename)

    If program = "FS Negative" Then
      'adding memo picture to schedule
        Windows(tempfilename).Activate
        Sheets(Sh.Name).Select
        ActiveSheet.Shapes("Picture MemoN").Select
        Selection.OnAction = "'" & tempfilename & "'!ShowSelectForms"
        ActiveSheet.Shapes("Supr_Appr").Select
        Selection.OnAction = "'" & tempfilename & "'!UserNameWindows"
        ActiveSheet.Shapes("Cler_Ent").Select
        Selection.OnAction = "'" & tempfilename & "'!ClericalApproval"
        ActiveSheet.Shapes("Edit Check FS Neg").Select
        Selection.OnAction = "'" & tempfilename & "'!snap_edit_check_neg"
        ActiveSheet.Shapes("SNAP_NegReturn").Select
        Selection.OnAction = "'" & tempfilename & "'!snap_neg_return"
        'ActiveSheet.Shapes("Return_SNAPPos").Select
        'Selection.OnAction = "'" & tempfilename & "'!snap_return"
        'ActiveSheet.Shapes("Return_TANF").Select
        'Selection.OnAction = "'" & tempfilename & "'!tanf_return"
        'ActiveSheet.Shapes("Return_MANeg").Select
        'Selection.OnAction = "'" & tempfilename & "'!MA_neg_return"
        'ActiveSheet.Shapes("Return_MA").Select
        'Selection.OnAction = "'" & tempfilename & "'!MA_return"
        ActiveSheet.Range("A1").Select

        
        'Adding Name Ranges to Cells after schedule is created
        Names.Add Name:="Client_Name", RefersTo:="=" & Sh.Name & "!$C$8"
        Names.Add Name:="Client_1stAddress", RefersTo:="=" & Sh.Name & "!$C$12"
        Names.Add Name:="Client_TownStateZip", RefersTo:="=" & Sh.Name & "!$C$13"
        Names.Add Name:="Sample_Month", RefersTo:="=" & Sh.Name & "!$AF$20"
        Names.Add Name:="Sample_Year", RefersTo:="=" & Sh.Name & "!$AI$20"
        Names.Add Name:="Examiner_Number", RefersTo:="=" & Sh.Name & "!$W$17"
        Names.Add Name:="Local_Agency", RefersTo:="=" & Sh.Name & "!$AA$20"
        Names.Add Name:="QC_Review_Number", RefersTo:="=" & Sh.Name & "!$C$20"
        Names.Add Name:="Case_Record_Number", RefersTo:="=" & Sh.Name & "!$L$20"
        Names.Add Name:="District_Code1", RefersTo:="=" & Sh.Name & "!$AD$20"
        Names.Add Name:="District_Code2", RefersTo:="=" & Sh.Name & "!$AE$20"
        Names.Add Name:="Taxonomy", RefersTo:="=" & Sh.Name & "!$E$50:$H$50"
        Names.Add Name:="County", RefersTo:="=" & Sh.Name & "!$M$5"
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    ElseIf program = "FS Positive" Or program = "FS Supplemental" Then
        CopyModule Workbooks("populate.xlsm"), "Module11", Workbooks(tempfilename)
        CopyModule Workbooks("populate.xlsm"), "ThisWorkbook", Workbooks(tempfilename)
        CopyModule Workbooks("populate.xlsm"), "Drop", Workbooks(tempfilename)
     
        
    
    'adding memo picture to schedule
        Windows(tempfilename).Activate
        Sheets(Sh.Name).Select
        ActiveSheet.Shapes("Picture Memo").Select
        Selection.OnAction = "'" & tempfilename & "'!ShowSelectForms"
        Range("a10").Select
        ActiveSheet.Shapes("Supr_Appr").Select
        Selection.OnAction = "'" & tempfilename & "'!UserNameWindows"
        ActiveSheet.Shapes("Cler_Ent").Select
        Selection.OnAction = "'" & tempfilename & "'!ClericalApproval"
        ActiveSheet.Shapes("Edit Check FS Pos").Select
        Selection.OnAction = "'" & tempfilename & "'!snap_edit_check_pos"
        ActiveSheet.Shapes("Return_SNAPPos").Select
        Selection.OnAction = "'" & tempfilename & "'!snap_return"
        ActiveSheet.Range("A1").Select
          
        'Adding Name Ranges to Cells after schedule is created
        Names.Add Name:="Client_Name", RefersTo:="=" & Sh.Name & "!$B$4"
        Names.Add Name:="Client_1stAddress", RefersTo:="=" & Sh.Name & "!$B$5"
        Names.Add Name:="Client_2ndAddress", RefersTo:="=" & Sh.Name & "!$B$6"
        Names.Add Name:="Client_TownStateZip", RefersTo:="=" & Sh.Name & "!$B$7"
        Names.Add Name:="Sample_Month", RefersTo:="=" & Sh.Name & "!$AD$18"
        Names.Add Name:="Sample_Year", RefersTo:="=" & Sh.Name & "!$AG$18"
        Names.Add Name:="Examiner_Number1", RefersTo:="=" & Sh.Name & "!$AJ$5"
        Names.Add Name:="Examiner_Number2", RefersTo:="=" & Sh.Name & "!$AK$5"
        Names.Add Name:="Local_Agency", RefersTo:="=" & Sh.Name & "!$X$18"
        Names.Add Name:="QC_Review_Number", RefersTo:="=" & Sh.Name & "!$A$18"
        Names.Add Name:="Case_Record_Number", RefersTo:="=" & Sh.Name & "!$I$18"
        Names.Add Name:="District_Code1", RefersTo:="=" & Sh.Name & "!$B$153"
        Names.Add Name:="District_Code2", RefersTo:="=" & Sh.Name & "!$C$153"
        Names.Add Name:="Taxonomy", RefersTo:="=" & Sh.Name & "!$C$159:$F$159"
        Names.Add Name:="County", RefersTo:="=" & Sh.Name & "!$M$5"
            
    Set sourceBook = ThisWorkbook
    Set destinationBook = Workbooks(tempfilename)
 
    With sourceBook.VBProject.VBComponents("ThisWorkbook").CodeModule
        codeFromSource = .Lines(1, .CountOfLines)
    End With
 

    Set DestCom = destinationBook.VBProject.VBComponents("ThisWorkbook")
    Set DestMod = DestCom.CodeModule
 
    With DestMod
        .DeleteLines 1, .CountOfLines
        .AddFromString codeFromSource
    End With
            Sheets(Sh.Name).Select
          '  ActiveSheet.Shapes("Drop Button").Select
          '  Selection.OnAction = "'" & tempfilename & "'!Drop.Drop"
            ActiveSheet.Shapes("Picture Memo").Select
            Selection.OnAction = "'" & tempfilename & "'!ShowSelectForms"
            Range("A1").Select
            Workbooks(tempfilename).Worksheets("FS Computation").Select
            ActiveSheet.Shapes("Butt1").Select
            Selection.OnAction = "'" & tempfilename & "'!Macro2"
            ActiveSheet.Shapes("Butt2").Select
            Selection.OnAction = "'" & tempfilename & "'!finalresults"
            ActiveSheet.Shapes("Butt3").Select
            Selection.OnAction = "'" & tempfilename & "'!SelfEmployment"
            ActiveSheet.Shapes("Butt4").Select
            Selection.OnAction = "'" & tempfilename & "'!moveselfemp"
            
            Workbooks(tempfilename).Worksheets("FS Workbook").Select
            ActiveSheet.Shapes("Supr_Appr").Select
            Selection.OnAction = "'" & tempfilename & "'!SuprWorkbook"
            ActiveSheet.Shapes("B 36").Select
            Selection.OnAction = "'" & tempfilename & "'!ClearButtons59to60"
            ActiveSheet.Shapes("B 174").Select
            Selection.OnAction = "'" & tempfilename & "'!ClearButtons89to135"
            ActiveSheet.Shapes("B 175").Select
            Selection.OnAction = "'" & tempfilename & "'!ClearButtons154to274"
            ActiveSheet.Shapes("B 176").Select
            Selection.OnAction = "'" & tempfilename & "'!ClearButtons220to279"
            ActiveSheet.Shapes("B 325").Select
            Selection.OnAction = "'" & tempfilename & "'!ClearButtons468to583"
            ActiveSheet.Shapes("B 326").Select
            Selection.OnAction = "'" & tempfilename & "'!ClearButtons548to555"
            ActiveSheet.Shapes("B 327").Select
            Selection.OnAction = "'" & tempfilename & "'!ClearButtons514to561"
            ActiveSheet.Shapes("B 362").Select
            Selection.OnAction = "'" & tempfilename & "'!ClearButtons635to696"
            ActiveSheet.Shapes("B 389").Select
            Selection.OnAction = "'" & tempfilename & "'!ClearButtons708to789"
            ActiveSheet.Shapes("B 500").Select
            Selection.OnAction = "'" & tempfilename & "'!ClearButtons802to849"
            ActiveSheet.Shapes("B 530").Select
            Selection.OnAction = "'" & tempfilename & "'!ClearButtons875to1023"
            ActiveSheet.Shapes("B 536").Select
            Selection.OnAction = "'" & tempfilename & "'!ClearButtons899to982"
            ActiveSheet.Shapes("B 604").Select
            Selection.OnAction = "'" & tempfilename & "'!ClearButtons932to968"
            ActiveSheet.Shapes("Cat El1").Select
            Selection.OnAction = "'" & tempfilename & "'!check_box_430_click"
            ActiveSheet.Shapes("CB None 311").Select
            Selection.OnAction = "'" & tempfilename & "'!Earned_Income_none"
            ActiveSheet.Shapes("CB None 331").Select
            Selection.OnAction = "'" & tempfilename & "'!Unearned_Income_none"
            ActiveSheet.Shapes("CB No Resources").Select
            Selection.OnAction = "'" & tempfilename & "'!No_Resources"
         '   ActiveSheet.Shapes("CB Comp 1 Only").Select
         '   Selection.OnAction = "'" & tempfilename & "'!Comp_1_Only_111"
         '   Selection.OnAction = "'" & tempfilename & "'!Comp_1_Only_311"
         '   Selection.OnAction = "'" & tempfilename & "'!Comp_1_Only_312"
         '   Selection.OnAction = "'" & tempfilename & "'!Comp_1_Only_313"
         '   Selection.OnAction = "'" & tempfilename & "'!Comp_1_Only_314"
         '   Selection.OnAction = "'" & tempfilename & "'!Comp_1_Only_323"
         '   Selection.OnAction = "'" & tempfilename & "'!Comp_1_Only_331"
         '   Selection.OnAction = "'" & tempfilename & "'!Comp_1_Only_332"
         '   Selection.OnAction = "'" & tempfilename & "'!Comp_1_Only_333"
         '   Selection.OnAction = "'" & tempfilename & "'!Comp_1_Only_334"
         '   Selection.OnAction = "'" & tempfilename & "'!Comp_1_Only_335"
         '   Selection.OnAction = "'" & tempfilename & "'!Comp_1_Only_336"
         '   Selection.OnAction = "'" & tempfilename & "'!Comp_1_Only_365"
         '   Selection.OnAction = "'" & tempfilename & "'!Comp_1_Only_366"
            
            'ActiveSheet.Shapes("CB 148").Select
            'Selection.OnAction = "'" & TempFileName & "'!check_box_80_click"
            'ActiveSheet.Shapes("CB 1112").Select
            'Selection.OnAction = "'" & TempFileName & "'!check_box_80_click"
         '   Workbooks(TempFileName).Worksheets("Survey").Select
         '   ActiveSheet.Shapes("OB 23").Select
         '   Selection.OnAction = "'" & TempFileName & "'!OB23_Click"
         '   ActiveSheet.Shapes("OB 24").Select
         '   Selection.OnAction = "'" & TempFileName & "'!OB24_Click"
         '   ActiveSheet.Shapes("OB 131").Select
         '   Selection.OnAction = "'" & TempFileName & "'!OB131_Click"
            Range("A1").Select
            
            ' Define Names
            
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    ElseIf program = "GA" Then
        CopyForm Workbooks("populate.xlsm"), "GAUserForm1", Workbooks(tempfilename)
        CopyForm Workbooks("populate.xlsm"), "GAUserForm2", Workbooks(tempfilename)
        CopyModule Workbooks("populate.xlsm"), "GAGetElements", Workbooks(tempfilename)
            Workbooks(tempfilename).Worksheets("GA Computation").Select
            ActiveSheet.Shapes("Butt1").Select
            Selection.OnAction = "'" & tempfilename & "'!GAClear"
            ActiveSheet.Shapes("Butt2").Select
            Selection.OnAction = "'" & tempfilename & "'!Module1.GAcompsheet"
            ActiveSheet.Shapes("Butt3").Select
            Selection.OnAction = "'" & tempfilename & "'!Compute"
            ActiveSheet.Shapes("Butt4").Select
            Selection.OnAction = "'" & tempfilename & "'!GAshow_form1"
            ActiveSheet.Shapes("Butt5").Select
            Selection.OnAction = "'" & tempfilename & "'!Module1.GAcompsheet"
            ActiveSheet.Shapes("Butt8").Select
            Selection.OnAction = "'" & tempfilename & "'!GAshow_form2"
             
            Workbooks(tempfilename).Worksheets("GA Workbook").Select
            ActiveSheet.Shapes("Supr_Appr").Select
            Selection.OnAction = "'" & tempfilename & "'!SuprWorkbook"
            'ActiveSheet.Shapes("GAClear Col. 2").Select
            'Selection.OnAction = "'" & tempfilename & "'!GAGetElements.Clearleft"
            'ActiveSheet.Shapes("GA Get Elements").Select
            'Selection.OnAction = "'" & tempfilename & "'!GAGetElements.GAgetresults"
            'ActiveSheet.Shapes("GAClear Col. 3").Select
            'Selection.OnAction = "'" & tempfilename & "'!GAGetElements.Clearright"
        'adding memo picture to schedule
        
        
        Windows(tempfilename).Activate
        Sheets(Sh.Name).Select
        ActiveSheet.Shapes("Picture Memo").Select
        Selection.OnAction = "'" & tempfilename & "'!ShowSelectForms"
        ActiveSheet.Shapes("Supr_Appr").Select
        Selection.OnAction = "'" & tempfilename & "'!UserNameWindows"
        ActiveSheet.Shapes("Cler_Ent").Select
        Selection.OnAction = "'" & tempfilename & "'!ClericalApproval"
        ActiveSheet.Range("A10").Select
        ActiveSheet.Shapes("Edit Check GA").OnAction = "'" & tempfilename & "'!ga_edit_check"
          
        'Adding Name Ranges to Cells after schedule is created
        Names.Add Name:="Client_Name", RefersTo:="=" & Sh.Name & "!$B$2"
        Names.Add Name:="Client_1stAddress", RefersTo:="=" & Sh.Name & "!$B$3"
        Names.Add Name:="Client_2ndAddress", RefersTo:="=" & Sh.Name & "!$B$4"
        Names.Add Name:="Client_TownStateZip", RefersTo:="=" & Sh.Name & "!$B$5"
        Names.Add Name:="SampleMonthYear", RefersTo:="=" & Sh.Name & "!$AB$10"
        Names.Add Name:="Examiner_Number1", RefersTo:="=" & Sh.Name & "!$AO$3"
        Names.Add Name:="Examiner_Number2", RefersTo:="=" & Sh.Name & "!$AP$3"
        Names.Add Name:="Local_Agency", RefersTo:="=" & Sh.Name & "!$U$10"
        Names.Add Name:="QC_Review_Number", RefersTo:="=" & Sh.Name & "!$A$10"
        Names.Add Name:="Case_Record_Number", RefersTo:="=" & Sh.Name & "!$I$10"
        Names.Add Name:="District_Code", RefersTo:="=" & Sh.Name & "!$Y$10"
        Names.Add Name:="Taxonomy", RefersTo:="=" & Sh.Name & "!$AG$92:$AJ$92"
        Names.Add Name:="County", RefersTo:="=" & Sh.Name & "!$O$4"
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
     ElseIf program = "TANF" Then
        CopyForm Workbooks("populate.xlsm"), "UserForm1", Workbooks(tempfilename)
        CopyForm Workbooks("populate.xlsm"), "UserForm2", Workbooks(tempfilename)
        CopyModule Workbooks("populate.xlsm"), "Drop", Workbooks(tempfilename)
            Workbooks(tempfilename).Worksheets("TANF Computation").Select
            ActiveSheet.Shapes("Butt4").OnAction = "'" & tempfilename & "'!TANFClear"
            ActiveSheet.Shapes("Butt3").OnAction = "'" & tempfilename & "'!Tanf"
            ActiveSheet.Shapes("Butt7").OnAction = "'" & tempfilename & "'!TANFCompute"
            ActiveSheet.Shapes("Butt8").OnAction = "'" & tempfilename & "'!show_form1"
            ActiveSheet.Shapes("Butt5").OnAction = "'" & tempfilename & "'!Tanf"
            ActiveSheet.Shapes("Butt9").OnAction = "'" & tempfilename & "'!show_form2"
            'ActiveSheet.Shapes("Return_TANF").Select
            'Selection.OnAction = "'" & tempfilename & "'!tanf_return"
           
            
            Workbooks(tempfilename).Worksheets("TANF Workbook").Select
            ActiveSheet.Shapes("Supr_Appr").Select
            Selection.OnAction = "'" & tempfilename & "'!SuprWorkbook"
            ActiveSheet.Shapes("B 536").OnAction = "'" & tempfilename & "'!ClearButtons88to90"
            ActiveSheet.Shapes("B 212").OnAction = "'" & tempfilename & "'!ClearButtons192to194"
            ActiveSheet.Shapes("B 328").OnAction = "'" & tempfilename & "'!ClearButtons207to300"
            ActiveSheet.Shapes("B 445").OnAction = "'" & tempfilename & "'!ClearButtons309to326"
            'Deactivated 2025-11-04: B 545 no longer exists in original template.
            'ActiveSheet.Shapes("B 545").OnAction = "'" & tempfilename & "'!ClearButtons375to390"
            
            'ActiveSheet.Shapes("B 634").OnAction = "'" & TempFileName & "'!ClearButtons493to495"
            'ActiveSheet.Shapes("B 719").OnAction = "'" & TempFileName & "'!ClearButtons500to539"

            Range("A1").Select

            
        'adding memo picture to schedule
        Windows(tempfilename).Activate
        Sheets(Sh.Name).Select
        ActiveSheet.Shapes("Drop Button").Select
        Selection.OnAction = "'" & tempfilename & "'!Drop.Drop"
        ActiveSheet.Shapes("Picture Memo").Select
        Selection.OnAction = "'" & tempfilename & "'!ShowSelectForms"
        ActiveSheet.Shapes("Supr_Appr").OnAction = "'" & tempfilename & "'!UserNameWindows"
        ActiveSheet.Shapes("Cler_Ent").OnAction = "'" & tempfilename & "'!ClericalApproval"
        ActiveSheet.Shapes("Edit Check TANF").OnAction = "'" & tempfilename & "'!tanf_edit_check"
        'ActiveSheet.Shapes("TANF_Return").Select
        'Selection.OnAction = "'" & tempfilename & "'!TANF_return"
        
        

        Range("a10").Select
          
        'Adding Name Ranges to Cells after schedule is created
        Names.Add Name:="Client_Name", RefersTo:="=" & Sh.Name & "!$B$2"
        Names.Add Name:="Client_1stAddress", RefersTo:="=" & Sh.Name & "!$B$3"
        Names.Add Name:="Client_TownStateZip", RefersTo:="=" & Sh.Name & "!$B$4"
        Names.Add Name:="SampleMonthYear", RefersTo:="=" & Sh.Name & "!$AB$10"
        Names.Add Name:="Examiner_Number1", RefersTo:="=" & Sh.Name & "!$AO$3"
        Names.Add Name:="Examiner_Number2", RefersTo:="=" & Sh.Name & "!$AP$3"
        Names.Add Name:="Local_Agency", RefersTo:="=" & Sh.Name & "!$U$10"
        Names.Add Name:="QC_Review_Number", RefersTo:="=" & Sh.Name & "!$A$10"
        Names.Add Name:="Case_Record_Number", RefersTo:="=" & Sh.Name & "!$I$10"
        Names.Add Name:="District_Code", RefersTo:="=" & Sh.Name & "!$Y$10"
        Names.Add Name:="Taxonomy", RefersTo:="=" & Sh.Name & "!$AF$85:$AI$85"
        Names.Add Name:="County", RefersTo:="=" & Sh.Name & "!$O$4"
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
       ElseIf program = "MA Positive" Then
 
    CopyForm Workbooks("populate.xlsm"), "MASelectForms", Workbooks(tempfilename)
    CopyForm Workbooks("populate.xlsm"), "UserFormMAC2", Workbooks(tempfilename)
    CopyForm Workbooks("populate.xlsm"), "UserFormMAC3", Workbooks(tempfilename)
    CopyModule Workbooks("populate.xlsm"), "Drop", Workbooks(tempfilename)
    CopyModule Workbooks("populate.xlsm"), "MA_Comp_mod", Workbooks(tempfilename)
    
    Set sourceBook = ThisWorkbook
    Set destinationBook = Workbooks(tempfilename)
 
    With sourceBook.VBProject.VBComponents("ThisWorkbook").CodeModule
        codeFromSource = .Lines(1, .CountOfLines)
    End With
 
'    With destinationBook.VBProject.VBComponents("ThisWorkbook").CodeModule
    Set DestCom = destinationBook.VBProject.VBComponents("ThisWorkbook")
    Set DestMod = DestCom.CodeModule
 
    With DestMod
        .DeleteLines 1, .CountOfLines
        .AddFromString codeFromSource
    End With
    
        'MA Tracking sheet macro assign buttons
        Windows(tempfilename).Activate
        Workbooks(tempfilename).Sheets("MA Tracking").Select
        ActiveSheet.Shapes("Trk_Supr_Appr").OnAction = "'" & tempfilename & "'!tracking_supr"
        
        'MA Workbook sheet macro assign buttons
        Workbooks(tempfilename).Sheets("MA Workbook").Select
        'ActiveSheet.Shapes("B 1").OnAction = "'" & tempfilename & "'!MAClearButtons1to15"
        'ActiveSheet.Shapes("B 2").OnAction = "'" & tempfilename & "'!MAClearButtons16to24"
        'ActiveSheet.Shapes("B 3").OnAction = "'" & tempfilename & "'!MAClearButtons25to27"
        'ActiveSheet.Shapes("B 4").OnAction = "'" & tempfilename & "'!MAClearButtons28to33"
        'ActiveSheet.Shapes("B 5").OnAction = "'" & tempfilename & "'!MAClearButtons34to45"
        'ActiveSheet.Shapes("B 6").OnAction = "'" & tempfilename & "'!MAClearButtons46to57"
        'ActiveSheet.Shapes("B 7").OnAction = "'" & tempfilename & "'!MAClearButtons58to66"
        'ActiveSheet.Shapes("B 8").OnAction = "'" & tempfilename & "'!MAClearButtons67to78"
        'ActiveSheet.Shapes("B 9").OnAction = "'" & tempfilename & "'!MAClearButtons79to93"
        'ActiveSheet.Shapes("B 10").OnAction = "'" & tempfilename & "'!MAClearButtons94to105"
        'ActiveSheet.Shapes("B 11").OnAction = "'" & tempfilename & "'!MAClearButtons106to117"
        'ActiveSheet.Shapes("B 12").OnAction = "'" & tempfilename & "'!MAClearButtons118to126"
        'ActiveSheet.Shapes("B 13").OnAction = "'" & tempfilename & "'!MAClearButtons127to132"
        ActiveSheet.Shapes("CB 354").OnAction = "'" & tempfilename & "'!resources_ex"
'        ActiveSheet.Shapes("Return_MA").Select
 '       Selection.OnAction = "'" & tempfilename & "'!MA_return"
        ActiveSheet.Range("A1").Select
        
        'MA Resources sheet macro assign buttons
        Workbooks(tempfilename).Sheets("MA Resources").Select
        ActiveSheet.Shapes("Final_QC1").OnAction = "'" & tempfilename & "'!show_formMAC2"
        ActiveSheet.Shapes("Final_QC2").OnAction = "'" & tempfilename & "'!show_formMAC3"
        ActiveSheet.Range("A1").Select
        
        'MA Income Comp sheet macro assign buttons
        Workbooks(tempfilename).Sheets("MA Income Comp").Select
        ActiveSheet.Shapes("Workbook_Transfer_Button").OnAction = "'" & tempfilename & "'!MA_Comp_Transfer_Workbook"
        ActiveSheet.Shapes("Income_Clear_Button").OnAction = "'" & tempfilename & "'!MA_Comp_Clear_Income"
        ActiveSheet.Shapes("SectionA_Transfer_Button").OnAction = "'" & tempfilename & "'!MA_Comp_Transfer_SectionA"
        ActiveSheet.Shapes("Clear_Wages_Button").OnAction = "'" & tempfilename & "'!MA_Comp_Clear_Wages"
        ActiveSheet.Range("A1").Select
        
        'PA 721 macro assign buttons
       ' Workbooks(tempfilename).Worksheets("PA721").Select
       ' ActiveSheet.Shapes("PA721_Fill_Button").Select
       ' Selection.OnAction = "'" & tempfilename & "'!PA721_Fill"
       ' ActiveSheet.Range("A1").Select
        
        '3rd party macros assign buttons
      '  Workbooks(tempfilename).Worksheets("3rd party").Select
      '  ActiveSheet.Shapes("TPLF_Button").Select
      '  Selection.OnAction = "'" & tempfilename & "'!TPL_Fill"
      '  ActiveSheet.Shapes("TPLF_Clear").Select
      '  Selection.OnAction = "'" & tempfilename & "'!TPL_Clear"
      '  ActiveSheet.Range("A1").Select
        
        'Survey worksheet assign buttons
     '   Workbooks(TempFileName).Worksheets("Survey").Select
     '   ActiveSheet.Shapes("OB 23").Select
     '   Selection.OnAction = "'" & TempFileName & "'!OB23_Click"
     '   ActiveSheet.Shapes("OB 24").Select
     '   Selection.OnAction = "'" & TempFileName & "'!OB24_Click"
     '   ActiveSheet.Shapes("OB 131").Select
     '   Selection.OnAction = "'" & TempFileName & "'!OB131_Click"
     '   ActiveSheet.Range("A1").Select
            
        'Main schedule assign buttons
        Sheets(Sh.Name).Select
        ActiveSheet.Shapes("Drop Button").OnAction = "'" & tempfilename & "'!Drop.Drop"
        ActiveSheet.Shapes("Clean_Case").OnAction = "'" & tempfilename & "'!ScheduleClean2"
        ActiveSheet.Shapes("Picture Memo").OnAction = "'" & tempfilename & "'!ShowSelectForms"
        ActiveSheet.Shapes("Supr_Appr").OnAction = "'" & tempfilename & "'!UserNameWindows"
        ActiveSheet.Shapes("Cler_Ent").OnAction = "'" & tempfilename & "'!ClericalApproval"
        ActiveSheet.Shapes("Edit Check").OnAction = "'" & tempfilename & "'!ma_edit_check_pos"
        ActiveSheet.Range("A10").Select
        
        'add name ranges
        Names.Add Name:="Client_Name", RefersTo:="=" & Sh.Name & "!$B$2"
        Names.Add Name:="Client_1stAddress", RefersTo:="=" & Sh.Name & "!$B$3"
        Names.Add Name:="Client_2ndAddress", RefersTo:="=" & Sh.Name & "!$B$4"
        Names.Add Name:="Client_TownStateZip", RefersTo:="=" & Sh.Name & "!$B$5"
        Names.Add Name:="SampleMonthYear", RefersTo:="=" & Sh.Name & "!$AM$10"
        Names.Add Name:="Examiner_Number1", RefersTo:="=" & Sh.Name & "!$AO$3"
        Names.Add Name:="Examiner_Number2", RefersTo:="=" & Sh.Name & "!$AP$3"
        Names.Add Name:="Local_Agency", RefersTo:="=" & Sh.Name & "!$AC$10"
        Names.Add Name:="QC_Review_Number", RefersTo:="=" & Sh.Name & "!$A$10:$F$10"
        Names.Add Name:="Case_Record_Number", RefersTo:="=" & Sh.Name & "!$H$10"
        Names.Add Name:="District_Code", RefersTo:="=" & Sh.Name & "!$AH$10"
        Names.Add Name:="Taxonomy", RefersTo:="=" & Sh.Name & "!$H$118:$K$118"
        Names.Add Name:="County", RefersTo:="=" & Sh.Name & "!$O$4"
        
        
        '*********** this commented out section was where we created memos for MA Nursing Homes
        'Set thisws = ActiveSheet
        'sPath = ActiveWorkbook.Path
        
        'Copy template data source to active directory
        'tempsourcename = sPath & "\FM DS Temp.xlsx"
        'SrceFile = DLetter & "\Finding Memo\Finding Memo Data Source.xlsx"
        'DestFile = tempsourcename
        'FileCopy SrceFile, DestFile

        'Open spreadsheet where all the data is stored to populate findings memo
        'Workbooks.Open Filename:=tempsourcename, UpdateLinks:=False

        'Dim datasourcewb As Workbook, datasourcews As Worksheet

        'Set datasourcewb = ActiveWorkbook
        'Set datasourcews = ActiveSheet

        'Client Information
        'Range("B2") = StrConv(thisws.Range("B2"), vbProperCase)
        'reformat street address
        'If thisws.Range("B4") <> "" Then
        '    Range("Z2") = StrConv(thisws.Range("B3"), vbProperCase) & ", " & StrConv(thisws.Range("B4"), vbProperCase)
        'Else
        '    Range("Z2") = StrConv(thisws.Range("B3"), vbProperCase)
        'End If

        'Reformat city and state
        'temparry = Split(thisws.Range("B5"), ",")
        'TempStr = StrConv(Trim(temparry(LBound(temparry))), vbProperCase)

        'Range("AA2") = TempStr & ", " & Trim(temparry(UBound(temparry)))

        'Case Number
        'Range("C2") = Right(thisws.Range("AC10"), 2) & "/" & Left(thisws.Range("H10"), 9)

        'review number
        'Range("D2") = thisws.Range("A10") & thisws.Range("B10") & thisws.Range("c10") & thisws.Range("D10") & thisws.Range("E10") & thisws.Range("F10")
        'review_number = Range("D2")
        'review month
        'Range("E2") = thisws.Range("AM10")
        'sample_month_separated = Right(thisws.Range("AM10"), 4) & Left(thisws.Range("AM10"), 2)
        'sample_month = Range("E2")
        'reviewer number
        'Range("A7") = Val(thisws.Range("AO3") & thisws.Range("AP3"))
        'supervisor/examiner
        'Range("G2") = Range("C7") & "/" & thisws.Range("AF2")
        'Range("R2") = thisws.Range("J61") & " - " & thisws.Range("O61")

        'Range("F7") = 2
        'Range("E7") = 2
        
        'datasourcewb.Close True
        
'        findmemo = "Nursing Home LRR Review Number " & review_number & _
'            " for Sample Month " & sample_month_separated & ".docx"
        
            'Copy template finding memo to active directory
'    SrceFile = strCurDir & "Nursing Home LRR.docx"
'    tempfindmemoname = sPath & "\FM temp.docx"
'    DestFile = tempfindmemoname
'    FileCopy SrceFile, DestFile
'
'    Set appWD = CreateObject("Word.Application")
'    appWD.Visible = True
'
'    With appWD
'    .Documents.Open Filename:=tempfindmemoname
'
'    With .ActiveDocument.MailMerge
'        .OpenDataSource Name:=(tempsourcename), _
'            ReadOnly:=True, LinkToSource:=0, AddToRecentFiles:=False, _
'            PasswordDocument:="", PasswordTemplate:="", WritePasswordDocument:="", _
'            WritePasswordTemplate:="", Revert:=False, Format:=0, _
'            Connection:="yourNamedRange", SQLStatement:="SELECT * FROM `yourNamedRange`", SQLStatement1:=""
'        .Destination = 0
'        .SuppressBlankLines = True
'        .Execute Pause:=False
'    End With
    
'Application.StatusBar = "Now creating document"
'
'    .ActiveDocument.Content.Font.Name = "Times New Roman"
'    '.ActiveDocument.Content.Font.Size = 12
'
'    .ActiveDocument.SaveAs sPath & "\" & findmemo
'    .ActiveDocument.Close
    
'    .Documents(tempfindmemoname).Close savechanges:=False

'End With
'appWD.Quit
'Set appWD = Nothing


 '       findmemo = "Nursing Home Business Office Review Number " & review_number & _
 '           " for Sample Month " & sample_month_separated & ".docx"
 '
            'Copy template finding memo to active directory
 '   SrceFile = strCurDir & "Nursing Home Business Office.docx"
 '   tempfindmemoname = sPath & "\FM temp.docx"
 '   DestFile = tempfindmemoname
 '   FileCopy SrceFile, DestFile
    
 '   Set appWD = CreateObject("Word.Application")
 '   appWD.Visible = True

 '   With appWD
 '   .Documents.Open Filename:=tempfindmemoname
 '
 '   With .ActiveDocument.MailMerge
 '       .OpenDataSource Name:=(tempsourcename), _
 '           ReadOnly:=True, LinkToSource:=0, AddToRecentFiles:=False, _
 '           PasswordDocument:="", PasswordTemplate:="", WritePasswordDocument:="", _
 '           WritePasswordTemplate:="", Revert:=False, Format:=0, _
 '           Connection:="yourNamedRange", SQLStatement:="SELECT * FROM `yourNamedRange`", SQLStatement1:=""
 '       .Destination = 0
 '       .SuppressBlankLines = True
 '       .Execute Pause:=False
 '   End With
    
'Application.StatusBar = "Now creating document"

'    .ActiveDocument.Content.Font.Name = "Times New Roman"
'    '.ActiveDocument.Content.Font.Size = 12

'    .ActiveDocument.SaveAs sPath & "\" & findmemo
'    .ActiveDocument.Close
    
'    .Documents(tempfindmemoname).Close savechanges:=False

'End With
'appWD.Quit
'Set appWD = Nothing



 '       findmemo = "MA Telephone Interview " & review_number & _
 '           " for Sample Month " & sample_month_separated & ".docx"
        
            'Copy template finding memo to active directory
 '   SrceFile = strCurDir & "MA Telephone Interview.docx"
 '   tempfindmemoname = sPath & "\FM temp.docx"
 '   DestFile = tempfindmemoname
 '   FileCopy SrceFile, DestFile
    
 '   Set appWD = CreateObject("Word.Application")
 '   appWD.Visible = True

'    With appWD
'    .Documents.Open Filename:=tempfindmemoname
    
'    With .ActiveDocument.MailMerge
'        .OpenDataSource Name:=(tempsourcename), _
'            ReadOnly:=True, LinkToSource:=0, AddToRecentFiles:=False, _
'            PasswordDocument:="", PasswordTemplate:="", WritePasswordDocument:="", _
'            WritePasswordTemplate:="", Revert:=False, Format:=0, _
'            Connection:="yourNamedRange", SQLStatement:="SELECT * FROM `yourNamedRange`", SQLStatement1:=""
'        .Destination = 0
'        .SuppressBlankLines = True
'        .Execute Pause:=False
'    End With
    
'Application.StatusBar = "Now creating document"

'    .ActiveDocument.Content.Font.Name = "Times New Roman"
    '.ActiveDocument.Content.Font.Size = 12

'    .ActiveDocument.SaveAs sPath & "\" & findmemo
'    .ActiveDocument.Close
    
'    .Documents(tempfindmemoname).Close savechanges:=False

'End With
'appWD.Quit
'Set appWD = Nothing

'Workbooks("FM DS Temp.xlsx").Close False

'delete temp files
'Kill tempfindmemoname
'Kill tempsourcename

        
      ElseIf program = "MA Closed" Or program = "MA CHIP Closed" Or program = "MA Rejected" Or program = "MA CHIP Rejected" Then
        Windows(tempfilename).Activate
        CopyForm Workbooks("populate.xlsm"), "MASelectForms", Workbooks(tempfilename)
        Sheets(Sh.Name).Select
        ActiveSheet.Shapes("Picture Memo").OnAction = "'" & tempfilename & "'!ShowSelectForms"
        ActiveSheet.Shapes("Supr_Appr").OnAction = "'" & tempfilename & "'!UserNameWindows"
        ActiveSheet.Shapes("Cler_Ent").OnAction = "'" & tempfilename & "'!ClericalApproval"
        ActiveSheet.Shapes("Edit Check").OnAction = "'" & tempfilename & "'!ma_edit_check_neg"
        'ActiveSheet.Shapes("Return_MANeg").Select
        'Selection.OnAction = "'" & tempfilename & "'!MA_neg_return"
        Range("a10").Select
     ElseIf program = "MA PE" Then
        Windows(tempfilename).Activate
        CopyForm Workbooks("populate.xlsm"), "MASelectForms", Workbooks(tempfilename)
        Sheets(Sh.Name).Select
        ActiveSheet.Shapes("Picture Memo").OnAction = "'" & tempfilename & "'!ShowSelectForms"
        ActiveSheet.Shapes("Supr_Appr").OnAction = "'" & tempfilename & "'!UserNameWindows"
        Range("a10").Select
        
    End If

            wb.SAVE
            
    If program = "FS Positive" Or program = "FS Supplemental" Then
             'get information from the new populate that includes information from BIS
            Populatenew
            wb.SAVE
    ElseIf program = "FS Negative" Then
            populate_snap_neg_delimited
            wb.SAVE
    ElseIf program = "TANF" Then
            Populate_TANF_Delimited
            wb.SAVE
  '  ElseIf program = "MA Positive" Then
  '          Populate_MA_Delimited
  '          wb.SAVE
    End If
    
            wb.Close (False)
            
           Application.DisplayAlerts = True
            
            Windows("populate.xlsm").Activate
            Sheets("Populate").Select
            
            For i = 2 To 40
                If Cells(i, 35) = Val(examnum) Then
                    For j = 38 To 77
                       If Cells(i, j) = "" Then
                            Cells(i, j) = CurDir & "\" & tempfilename
                        Exit For
                        End If
                    Next j
                   Exit For
               End If
           Next i
        End If
 Next Sh
 
If program = "FS Positive" Or program = "FS Supplemental" Or program = "FS Negative" Or program = "TANF" Then 'add back in MA Positive when ready
    'Close BIS spreadsheet without saving
    wb_bis.Close False
End If
 
 Workbooks(fileName).SAVE
 Workbooks(fileName).Close
 'Workbooks("FS Workbook.xlsm").Close
 

 Call send_emails(program, monthstr)
 

 
 End Sub
 
'Option Explicit 'Meant to enforce variable declarations, which helps type safety.
'Not necessary persay, but highly recommended to prevent programming errors.

' Although this function fits everything needed at the end of the macro,
' functions should generally only have a single responsibility, for ease of maintenance.
' Recommend that function should split in three: ProcessExaminerFiles for file operations
' SendEmails for email creation and sending, and Cleanup for file and variable cleanup.

'Function to process and distribute reviews, send confirmation emails, and cleanup working directory.
Public Function send_emails(ByVal program As String, ByVal monthstr As String)

    ' --- Constants ---
    Const POPULATE_WB_NAME As String = "populate.xlsm"
    Const POPULATE_WS_NAME As String = "Populate"
    Const EMAIL_COL As String = "AK"    ' Column with email addresses
    Const NAME_COL As String = "AJ"     ' Column with examiner names
    Const ID_COL As String = "AI"       ' Column with examiner IDs/numbers
    Const FILES_START_COL As String = "AL"
    Const FILES_END_COL As String = "BY"
    ' pathdir is a global module variable used in this function.
    ' This is not best practice, and should be passed to the function explictly, or within the sheet.
    
    ' --- Excel Objects ---
    Dim wbPopulate As Workbook
    Dim wsPopulate As Worksheet
    Dim cell As Range
    Dim fileCell As Range
    Dim emailRng As Range
    Dim filesRng As Range

    ' --- Outlook Objects ---
    Dim OutApp As Object ' Outlook.Application
    Dim OutMail As Object ' Outlook.MailItem

    ' --- File System & Path Variables ---
    Dim reviewMonthFolderName As String
    Dim examinerFolderName As String
    Dim fullExaminerPath As String
    Dim reviewFilePath As String
    Dim reviewFileName As String
    Dim clientName As String
    Dim reviewNumber As String
    Dim reviewFolderPath As String
    Dim programFolder As String

    ' --- Helper Variables ---
    Dim examinerName As String
    Dim examinerNameParts As Variant
    Dim examinerLastName As String
    Dim examinerID As String
    Dim emailSubject As String
    Dim emailBody As String
    Dim monthText As String
    Dim monthNumber As String
    Dim yearPart As String
    Dim tempArray As Variant
    Dim fileExists As Boolean
    Dim folderExists As Boolean
    Dim displayPath As String ' For email body error checking.
    Dim stopMacro As Boolean ' Flag to indicate if macro should stop after error
    Dim filesToDelete As Object ' Collection to store the paths of all successfully copied files.

    ' --- Error Handling ---
    Dim errNum As Long
    Dim errDesc As String
    Dim contextMsg As String ' For detailed error messages

    ' --- Initialization ---
    On Error GoTo ErrorHandler ' General error handler
    stopMacro = False ' Initialize flag
    Set filesToDelete = CreateObject("Scripting.Dictionary") 'Use Dictionary for easier checks for additions.
    

    ' Get the Populate workbook and worksheet
    On Error Resume Next ' Temporarily ignore error if workbook isn't open
    Set wbPopulate = Workbooks(POPULATE_WB_NAME)
    On Error GoTo ErrorHandler ' Restore general error handler; this approach is used throughout to gather error info.
    If wbPopulate Is Nothing Then
        MsgBox POPULATE_WB_NAME & " is not open. Please open it and try again.", vbCritical
        Exit Function
    End If
    Set wsPopulate = wbPopulate.Sheets(POPULATE_WS_NAME)

    ' Set up Outlook
    On Error Resume Next ' Handle error if Outlook is not available
    Set OutApp = GetObject(, "Outlook.Application") ' Try getting existing instance
    If Err.Number <> 0 Then
        Set OutApp = CreateObject("Outlook.Application") ' Otherwise create new instance
    End If
    On Error GoTo ErrorHandler ' Restore general error handler
    If OutApp Is Nothing Then
        MsgBox "Could not create Outlook application object. Check if Outlook is installed.", vbCritical
        Exit Function
    End If

    Application.ScreenUpdating = False
    Application.StatusBar = "Processing emails... Please wait."

    ' --- Main Loop: Iterate through email addresses ---
    On Error Resume Next
    Set emailRng = wsPopulate.Columns(EMAIL_COL).SpecialCells(xlCellTypeConstants)
    On Error GoTo ErrorHandler
    If emailRng Is Nothing Then
        MsgBox "No email addresses found in column " & EMAIL_COL & ".", vbInformation
        GoTo CleanUp ' Skip to cleanup on missing emails.
    End If

    For Each cell In emailRng.Cells
        If stopMacro Then GoTo CleanUp ' Check if a previous error occurred

        ' Basic email format check
        If cell.Value Like "?*@?*.?*" Then
            ' Define the range for file paths in this row
            Set filesRng = wsPopulate.Range(wsPopulate.Cells(cell.Row, FILES_START_COL), wsPopulate.Cells(cell.Row, FILES_END_COL))

            ' Check if there are any file paths listed for this Examiner
            If Application.WorksheetFunction.CountA(filesRng) > 0 Then
                ' --- Get Examiner Info ---
                examinerName = wsPopulate.Cells(cell.Row, NAME_COL).Value
                examinerID = wsPopulate.Cells(cell.Row, ID_COL).Value

                ' --- Prepare Path and Folder Info ---
                yearPart = Left(monthstr, 4)
                monthNumber = Right(monthstr, 2)
                If IsNumeric(monthNumber) Then
                    monthText = monthName(CLng(monthNumber)) ' Use built-in date callbacks instead of large case statements.
                Else
                    monthText = "InvalidMonth"
                End If
                reviewMonthFolderName = "Review Month " & monthText & " " & yearPart

                ' --- Determine Program Folder Name ---
                programFolder = program 'Default
                Select Case program
                    Case "FS Supplemental"
                        programFolder = "FS Positive"
                    Case "MA Closed", "MA CHIP Closed", "MA Rejected", "MA CHIP Rejected"
                        programFolder = "MA Negative"
                        ' Add other cases if needed
                End Select

                ' --- Build Base Path for Examiner/Program/Month ---
                examinerNameParts = Split(Trim(examinerName), " ") ' Split full name by spaces
                examinerLastName = examinerNameParts(UBound(examinerNameParts)) ' Get the last element of all parts
                
                examinerFolderName = examinerLastName & " - " & examinerID
                fullExaminerPath = pathdir & examinerFolderName & "\" & programFolder & "\" & reviewMonthFolderName & "\"

                ' --- Create Base Examiner/Program/Month Folder ---
                contextMsg = "checking base folder: " & fullExaminerPath
                folderExists = False
                On Error Resume Next
                folderExists = (Dir(fullExaminerPath, vbDirectory) <> "")
                If Err.Number <> 0 Then GoTo FatalError ' Error during Dir check
                On Error GoTo ErrorHandler ' Restore

                If Not folderExists Then
                    contextMsg = "creating base folder: " & fullExaminerPath
                    On Error Resume Next
                    MkDir fullExaminerPath
                    If Err.Number <> 0 Then GoTo FatalError ' Error during MkDir
                    On Error GoTo ErrorHandler ' Restore
                End If
                ' --- Base folder should now exist ---

                ' --- Prepare Email Object ---
                Set OutMail = OutApp.CreateItem(0) ' olMailItem
                emailSubject = program & " Schedule Folder Created for " & monthText & " " & yearPart
                displayPath = Replace(fullExaminerPath, Left(pathdir, InStr(pathdir, "\")), "...") ' More robust relative path idea

                If program = "TANF CAR" Then
                    ' Added extra html formatting for clarity.
                    emailBody = "Hi " & examinerName & ",<br><br>" & _
                                "Your " & program & " folders for " & _
                                monthText & " " & yearPart & " were created in the folder: <br>" & _
                                "<i>" & displayPath & "</i>" & "<br><br>Thank you."
                Else
                    emailBody = "Hi " & examinerName & ",<br><br>" & _
                                "Your " & program & " Schedules for " & _
                                monthText & " " & yearPart & " have been copied to the folder: <br>" & _
                                "<i>" & displayPath & "</i>" & "<br><br>Thank you."
                End If

                ' --- Inner Loop: Process Files - Any error here stops the macro ---
                Dim processingErrorOccurred As Boolean
                processingErrorOccurred = False ' Reset for each examiner

                For Each fileCell In filesRng.SpecialCells(xlCellTypeConstants)
                    reviewFilePath = Trim(fileCell.Value)
                    fileExists = False
                    clientName = "" ' Reset client name
                    reviewNumber = "" ' Reset review number

                    If reviewFilePath <> "" Then
                        ' 1. Check if the source file exists in case of saving errors.
                        contextMsg = "checking source file: " & reviewFilePath
                        On Error Resume Next
                        fileExists = (Dir(reviewFilePath) <> "")
                        If Err.Number <> 0 Then GoTo FatalError ' Error during Dir check
                        On Error GoTo ErrorHandler
                        If Not fileExists Then GoTo FatalError ' Source file MUST exist

                        ' 2. Extract filename and review number
                        contextMsg = "parsing filename: " & reviewFilePath
                        On Error Resume Next ' Handle potential errors in Split (rare)
                        tempArray = Split(reviewFilePath, "\")
                        reviewFileName = tempArray(UBound(tempArray))

                        ' Extract review number, assuming "Review Number XXXXXX..."
                        tempArray = Split(reviewFileName, " ")
                        If UBound(tempArray) >= 2 Then
                            reviewNumber = tempArray(2)
                        Else
                            GoTo FatalError
                        End If
                        If Err.Number <> 0 Then GoTo FatalError ' Error during string parsing
                        On Error GoTo ErrorHandler

                        ' --- *** Performance Bottleneck *** ---
                        ' Opening each workbook is very slow just to get a single piece of data each.
                        ' Consider adding client names to worksheet or passing them to this function directly
                        ' during a prior step if possible.

                        ' 3. Open Workbook and Get Client Name
                        Dim wbReview As Workbook
                        Dim wsReview As Worksheet
                        contextMsg = "opening workbook: " & reviewFileName
                        On Error Resume Next
                        Set wbReview = Workbooks.Open(fileName:=reviewFilePath, UpdateLinks:=False, ReadOnly:=True)
                        If Err.Number <> 0 Then GoTo FatalError ' Error opening workbook
                        On Error GoTo ErrorHandler

                        contextMsg = "getting client name from " & reviewFileName & " (Program: " & program & ")"
                        On Error Resume Next ' Handle errors getting name
                        Select Case program
                            Case "FS Positive", "TANF", "GA", "FS Supplemental"
                                For Each wsReview In wbReview.Worksheets
                                    If Right(wsReview.Name, 11) = "Computation" Then clientName = wsReview.Range("B1").Value: Exit For
                                Next
                            Case "FS Negative"
                                ' Partially Explicit value call for extra safety
                                ' Recommend setting either an explicit sheet name like wbReview.Worksheets("Sheet1")
                                '  or index like wbReview.Worksheets(1) if we know the order to avoid errors/issues.
                                clientName = wbReview.ActiveSheet.Range("C8").Value
                            Case "MA Closed", "MA CHIP Closed", "MA Rejected", "MA CHIP Rejected", "MA Negative"
                                clientName = wbReview.ActiveSheet.Range("B7").Value
                            Case "MA PE"
                                clientName = wbReview.ActiveSheet.Range("B7").Value
                            Case "MA Positive"
                                For Each wsReview In wbReview.Worksheets
                                    If wsReview.Name = "MA Workbook" Then clientName = wsReview.Range("C15").Value: Exit For
                                Next
                            Case "TANF CAR"
                                clientName = wbReview.ActiveSheet.Range("B2").Value
                            Case Else
                                clientName = "ProgNotFound"
                        End Select
                        If Err.Number <> 0 Then ' Error occurred during name extraction
                            wbReview.Close SaveChanges:=False ' Try to close workbook even if name extraction failed
                            GoTo FatalError
                        End If
                        On Error GoTo ErrorHandler ' Restore handler

                        wbReview.Close SaveChanges:=False
                        Set wbReview = Nothing
                        If Trim(clientName) = "" Or clientName = "ProgNotFound" Then
                            clientName = IIf(Trim(clientName) = "", "NameNotFound", clientName) ' Clarify if blank or program mismatch
                            GoTo FatalError ' Client name MUST be found
                        End If

                        ' 4. Create Review Sub-Folder
                        reviewFolderPath = fullExaminerPath & reviewNumber & " - " & clientName & "\"
                        contextMsg = "checking review sub-folder: " & reviewFolderPath
                        folderExists = False
                        On Error Resume Next
                        folderExists = (Dir(reviewFolderPath, vbDirectory) <> "")
                        If Err.Number <> 0 Then GoTo FatalError ' Error during Dir check
                        On Error GoTo ErrorHandler

                        If Not folderExists Then
                            contextMsg = "creating review sub-folder: " & reviewFolderPath
                            On Error Resume Next
                            MkDir reviewFolderPath
                            If Err.Number <> 0 Then GoTo FatalError ' Error during MkDir
                            On Error GoTo ErrorHandler
                        End If

                        ' 5. Copy File (if not TANF CAR)
                        If program <> "TANF CAR" Then
                            contextMsg = "copying file " & reviewFileName & " to " & reviewFolderPath
                            On Error Resume Next
                            FileCopy reviewFilePath, reviewFolderPath & reviewFileName
                            If Err.Number <> 0 Then GoTo FatalError ' Error during FileCopy
                            On Error GoTo ErrorHandler
                            
                            ' Track successful copies to delete at the end, avoiding another search.
                            If Not filesToDelete.Exists(reviewFilePath) Then
                                filesToDelete.Add reviewFilePath, reviewFilePath 'We add the path as both search key and stored item.
                            End If
                        End If

                        ' 6. Add MA Positive back if needed, mirroring previous error system.
                        'If MA positive, look for memos for NH
                        'If program = "MA Positive" Then
                        '
                        '     FilenameNH = sPath & "\Nursing Home LRR Review Number " & JustReviewNumber & " for Sample Month " & monthstr & ".docx"
                        '
                        '     If Len(Dir(FilenameNH)) > 0 Then
                        '
                        '     ' Find just the name of the file without the path
                        '         temparray = Split(FilenameNH, "\")
                        '         JustNHFileName = temparray(UBound(temparray))
                        '         FileCopy FilenameNH, ExtPathStr & JustNHFileName
                        '     End If
                        '
                        '     FilenameNH = sPath & "\Nursing Home Business Office Review Number " & JustReviewNumber & " for Sample Month " & monthstr & ".docx"
                        '
                        '     If Len(Dir(FilenameNH)) > 0 Then
                        '
                        '     ' Find just the name of the file without the path
                        '         temparray = Split(FilenameNH, "\")
                        '         JustNHFileName = temparray(UBound(temparray))
                        '         FileCopy FilenameNH, ExtPathStr & JustNHFileName
                        '     End If
                        '
                        '      FilenameNH = sPath & "\MA Telephone Interview " & JustReviewNumber & " for Sample Month " & monthstr & ".docx"
                        '
                        '      If Len(Dir(FilenameNH)) > 0 Then
                            
                        '     ' Find just the name of the file without the path
                        '         temparray = Split(FilenameNH, "\")
                        '         JustNHFileName = temparray(UBound(temparray))
                        '         FileCopy FilenameNH, ExtPathStr & JustNHFileName
                        '     End If
                        'End If

                    End If ' End If reviewFilePath <> ""
                Next fileCell ' Next file for this examiner

                ' --- Send/Display Email ONLY if NO errors occurred for this examiner ---
                If Not processingErrorOccurred Then
                    With OutMail
                        .To = cell.Value ' *** The Actual Fix, though the whole function needed TLC.***
                        ' .CC = "..." ' Add CC if needed
                        .Subject = emailSubject
                        .HTMLBody = emailBody
                        .Importance = 1 ' olImportanceNormal
                        .DeleteAfterSubmit = True
                        .NoAging = True ' Optional: Prevents auto-archiving, but included for tradition
                        .Send ' Or .Display to show to user for checking. Note that it will not send automatically if .Display is used.
                    End With
                End If

                Set OutMail = Nothing ' Release email object

            End If ' End If CountA(filesRng) > 0
        End If ' End If cell.Value Like "?*@?*.?*"
    Next cell ' Next examiner

CleanUp:
    ' --- Clean up generated files ---
    If filesToDelete.Count > 0 Then
        Application.StatusBar = "Cleaning up " & filesToDelete.Count & " generated files."
        Dim filePathToDelete As Variant
        Dim deleteErrors As String
        deleteErrors = ""
        contextMsg = "Deleting copied review files from working directory"
        
        For Each filePathToDelete In filesToDelete.Keys ' Iterate through the collected paths
            On Error Resume Next ' Make Kill non-fatal for this loop
            Kill CStr(filePathToDelete) ' Ensure we're feeding Kill a string.
            If Err.Number <> 0 Then
                deleteErrors = deleteErrors & vbCrLf & " - " & filePathToDelete & " (Error " & Err.Number & ")"
                Err.Clear
            End If
            ' No need to restore the error handler, as Resume Next continues the loop.
        Next filePathToDelete
        On Error GoTo ErrorHandler ' We restore the error handler after the loop completes.
        
        If Len(deleteErrors) > o Then
            MsgBox "File Processing completed successfully, but some working files were not deleted: " & deleteErrors, vbExclamation, "Cleanup Warning"
        End If
    End If 'End If filesToDelete.Count > 0


    ' --- Release Objects and Restore Settings ---
    If Not OutMail Is Nothing Then Set OutMail = Nothing ' Ensure release if loop exited early
    Set fileCell = Nothing
    Set filesRng = Nothing
    Set cell = Nothing
    Set emailRng = Nothing
    Set wsPopulate = Nothing
    Set wbPopulate = Nothing
    Set OutApp = Nothing
    Set filesToDelete = Nothing
    

    Application.StatusBar = False
    Application.ScreenUpdating = True

    If stopMacro Then
         MsgBox "Macro stopped due to the error reported above.", vbExclamation, "Macro Halted"
    Else
         MsgBox "Processing complete.", vbInformation, "Macro Finished"
         ' Added additional successful confirmation message after macro finish.
    End If

    Exit Function ' Normal or error exit

FatalError:
    ' --- Stop Macro Immediately and Report Specific Error ---
    errNum = Err.Number
    errDesc = Err.Description
    MsgBox "FATAL ERROR occurred while " & contextMsg & vbCrLf & vbCrLf & _
           "Error Number: " & errNum & vbCrLf & _
           "Description: " & errDesc & vbCrLf & vbCrLf & _
           "Macro execution will stop. Please fix the issue and retry. ", vbCritical, "Processing Error - Macro Stopping"
    stopMacro = True ' Set flag to prevent further processing/emails
    Stop 'Allows for debug before initiating cleanup
    On Error GoTo 0 ' Disable error handling before jumping
    GoTo CleanUp ' Jump to cleanup routine

ErrorHandler:
    ' --- Basic Error Reporting for Unexpected Errors ---
    errNum = Err.Number
    errDesc = Err.Description
    MsgBox "An UNEXPECTED error occurred:" & vbCrLf & _
           "Error Number: " & errNum & vbCrLf & _
           "Description: " & errDesc & vbCrLf & _
           "Context: " & contextMsg & vbCrLf & _
           "Procedure: send_emails_refactored_stopOnError" & vbCrLf & vbCrLf & _
           "Macro execution will stop. ", vbCritical, "Unexpected Error - Macro Stopping"
    stopMacro = True
    Stop
    On Error GoTo 0
    GoTo CleanUp ' Jump to cleanup routine

End Function


Sub Ga()
        Windows(fileName).Activate
        Sheets("GA").Name = Sheets("Temp").Range("A1")
         For i = 1 To maxrow
         strTemp = "Processing Form " & i & " of " & maxrow & ". Please be patient..."
         Application.StatusBar = strTemp
            If i > 1 Then
            TempStr = Sheets("Temp").Range("A" & i - 1)
            Sheets(TempStr).Select
            Sheets(TempStr).Copy After:=Sheets(TempStr)
            TempStr = TempStr & " (2)"
            Sheets(TempStr).Name = Sheets("Temp").Range("A" & i)
            End If
             'looking up district and bringing over district
          Range("Y10") = ""
          Select Case Sheets("Temp").Range("D" & i)
            Case 2
                TempStr2 = _
                "=VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R2C15:R9C16,2,FALSE)"
                Range("Y10").FormulaR1C1 = TempStr2
             Case 9
               TempStr2 = _
                "=VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R10C15:R11C16,2,FALSE)"
                Range("Y10").FormulaR1C1 = TempStr2
            Case 23
                TempStr2 = _
                "=VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R11C15:R13C16,2,FALSE)"
                Range("Y10").FormulaR1C1 = TempStr2
            Case 40
               TempStr2 = _
                "=VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R14C15:R15C16,2,FALSE)"
                Range("Y10").FormulaR1C1 = TempStr2
            Case 46
               TempStr2 = _
                "=VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R16C15:R17C16,2,FALSE)"
                Range("Y10").FormulaR1C1 = TempStr2
            Case 51
               TempStr2 = _
                "=VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R18C15:R36C16,2,FALSE)"
                Range("Y10").FormulaR1C1 = TempStr2
            Case 63
               TempStr2 = _
                "=VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R37C15:R38C16,2,FALSE)"
                Range("Y10").FormulaR1C1 = TempStr2
            Case 65
               TempStr2 = _
                "=VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R39C15:R41C16,2,FALSE)"
                Range("Y10").FormulaR1C1 = TempStr2
          End Select
          Range("Y10").Select
          Selection.Copy
          Range("Y10").Select
          Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
          'putting name in examiner field.  Looking up examiner number
          TempStr4 = _
          "=VLOOKUP(Temp!R" & i & "C14,[populate.xlsm]Populate!R2C27:R39C29,2,FALSE)"
                Range("AG2").FormulaR1C1 = TempStr4
          Range("AG2").Select
          Selection.Copy
          Range("AG2").Select
          Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
          If i = 1 Then
         'bringing over month
            monthstr1 = Application.WorksheetFunction.Text(revmonth, "MMYYYY")
            Range("AB10") = monthstr1
          End If
         'bringing over examiner number
            If (Len(Sheets("Temp").Range("N" & i))) < 2 Then
                Range("AO3") = "0"
            Else
             Range("AO3") = Left(Sheets("Temp").Range("N" & i), 1)
            End If
             Range("AP3") = Right(Sheets("Temp").Range("N" & i), 1)
         'determining if county has only one number or two numbers in spreadsheet
            If Sheets("Temp").Range("D" & i) > 9 Then
            countyunit = Sheets("Temp").Range("D" & i)
            Else
            countyunit = "0" & Sheets("Temp").Range("D" & i)
            End If
            Range("U10") = Sheets("Temp").Range("C" & i) & countyunit
         'putting county name
            TempStr = "=VLOOKUP(Temp!R" & i & "C4,[populate.xlsm]Populate!R2C30:R68C31,2,FALSE)"
            Range("O4").FormulaR1C1 = TempStr
            Range("O4").Select
            Selection.Copy
            Range("O4").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
              :=False, Transpose:=False
            Range("O4") = Range("O4") & " County"
         'district name
            Range("P5") = ""
            distnum = Range("Y10")
            If distnum <> "" Then
            If Val(distnum) < 11 Then
            TempStr2 = "=VLOOKUP(" & """" & distnum & """" & ",[populate.xlsm]Populate!R2C16:R41C18,3,FALSE)"
            Else
            TempStr2 = "=VLOOKUP(" & distnum & ",[populate.xlsm]Populate!R2C16:R41C18,3,FALSE)"
            End If
            Range("P5").FormulaR1C1 = TempStr2
            Range("P5").Select
            Selection.Copy
            Range("P5").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
              :=False, Transpose:=False
            Range("P5") = Range("P5") & " District"
            End If
         'concatenate county and district
            If Range("P5") <> "" Then
            Range("O4") = Range("O4") & " , " & Range("P5")
            Range("P5") = ""
            Else
            Range("O4") = Range("O4")
            End If
         'bringing over review number
            Range("A10") = Sheets("Temp").Range("A" & i)
         'bringing over case number
         TempStr = Sheets("Temp").Range("F" & i)
         strlen = Len(TempStr)
         If strlen < 7 Then
            For j = 1 To 7 - strlen
                TempStr = "0" & TempStr
            Next j
         End If
            Range("I10") = TempStr
         'bringing over case category
            Range("Q10") = Sheets("Temp").Range("G" & i)
         'bringing over first and last name
             Range("B2") = Sheets("Temp").Range("I" & i) & " " & Sheets("Temp").Range("H" & i)
         'bringing over Grant Group
             'Range("S10") = ""
               'If Sheets("Temp").Range("Y" & i) <> "0" Then
                 ' Range("S10") = Sheets("Temp").Range("Y" & i)
             ' End If
         'bringing over address
             Range("B3") = Sheets("Temp").Range("J" & i)
         'bringing over second address
             Range("B4") = ""
             If Sheets("Temp").Range("Z" & i) <> "0" Then
             Range("B4") = Sheets("Temp").Range("Z" & i)
             End If
         'bringing over city, state, and zip
             Range("B5") = Sheets("Temp").Range("K" & i) & ",  " & Sheets("Temp").Range("L" & i) _
             & " " & Sheets("Temp").Range("M" & i)
            Range("A1").Select
         Next i
End Sub
Sub Mapos()
          Sheets("MA Pos").Name = Sheets("Temp").Range("A1")
         For i = 1 To maxrow
         strTemp = "Processing Form " & i & " of " & maxrow & ". Please be patient..."
         Application.StatusBar = strTemp
         ' Save, close and open the schedule workbook after 50 worksheets are added.
         ' This prevents the copy/paste operation from hanging
            If i Mod 50 = 0 Then
            
                Application.DisplayAlerts = False
                Windows(fileName).Activate
                ActiveWorkbook.SAVE
                Windows("populate.xlsm").Activate
                Windows(fileName).Close (False)
                Workbooks.Open fileName, UpdateLinks:=False
                Application.DisplayAlerts = True
                 
            End If
            
           If Range("CD1") = "" Then
            Range("CD1") = Format(Date, "short date")
          End If
            
            If i > 1 Then
            TempStr = Sheets("Temp").Range("A" & i - 1)
            Sheets(TempStr).Select
            Sheets(TempStr).Copy After:=Sheets(TempStr)
            TempStr = TempStr & " (2)"
            Sheets(TempStr).Name = Sheets("Temp").Range("A" & i)
            End If
          'looking up district and bringing over district
          Range("AH10") = ""
         Select Case Sheets("Temp").Range("D" & i)
            Case 2
                TempStr2 = _
                "=VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R2C15:R9C16,2,FALSE)"
                Range("AH10").FormulaR1C1 = TempStr2
  '           Case 9
  '             TempStr2 = _
  '              "=VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R10C15:R11C16,2,FALSE)"
  '              Range("AH10").FormulaR1C1 = TempStr2
            Case 23
                TempStr2 = _
                "=VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R11C15:R13C16,2,FALSE)"
                Range("AH10").FormulaR1C1 = TempStr2
            Case 40
               TempStr2 = _
                "=VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R14C15:R15C16,2,FALSE)"
                Range("AH10").FormulaR1C1 = TempStr2
  '          Case 46
  '             TempStr2 = _
  '              "=VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R16C15:R17C16,2,FALSE)"
  '              Range("AH10").FormulaR1C1 = TempStr2
            Case 51
               TempStr2 = _
                "=VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R18C15:R36C16,2,FALSE)"
                Range("AH10").FormulaR1C1 = TempStr2
            Case 63
               TempStr2 = _
                "=VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R37C15:R38C16,2,FALSE)"
                Range("AH10").FormulaR1C1 = TempStr2
            Case 65
               TempStr2 = _
                "=VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R39C15:R41C16,2,FALSE)"
                Range("AH10").FormulaR1C1 = TempStr2
          End Select
          Range("AH10").Select
          Selection.Copy
          Range("AH10").Select
          Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
          'putting name in examiner field.  Looking up examiner number
          TempStr4 = _
          "=VLOOKUP(Temp!R" & i & "C14,[populate.xlsm]Populate!R2C27:R39C28,2,FALSE)"
                Range("AF2").FormulaR1C1 = TempStr4
          Range("AF2").Select
          Selection.Copy
          Range("AF2").Select
          Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
          If i = 1 Then
         'bringing over month
            monthstr1 = Application.WorksheetFunction.Text(revmonth, "MM/YYYY")
            Range("AM10") = monthstr1
          End If
         'bringing over examiner number
           If (Len(Sheets("Temp").Range("N" & i))) < 2 Then
                Range("AO3") = "0"
             Else
              Range("AO3") = Left(Sheets("Temp").Range("N" & i), 1)
             End If
            Range("AP3") = Right(Sheets("Temp").Range("N" & i), 1)
         'bringing over area
            Range("AC10") = Sheets("Temp").Range("C" & i)
         'determining if county has only one number or two numbers in spreadsheet
           
            If Sheets("Temp").Range("D" & i) > 9 Then
            countyunit = Sheets("Temp").Range("D" & i)
            Else
            countyunit = "0" & Sheets("Temp").Range("D" & i)
            End If
            Range("AC10") = Sheets("Temp").Range("C" & i) & countyunit
        'putting county name
            TempStr = "=VLOOKUP(Temp!R" & i & "C4,[populate.xlsm]Populate!R2C30:R68C31,2,FALSE)"
            Range("O4").FormulaR1C1 = TempStr
            Range("O4").Select
            Selection.Copy
            Range("O4").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
              :=False, Transpose:=False
            Range("O4") = Range("O4") & " County"
        'district name
            Range("P5") = ""
            distnum = Range("AH10")
            If distnum <> "" Then
            If Val(distnum) < 11 Then
            TempStr2 = "=VLOOKUP(" & """" & distnum & """" & ",[populate.xlsm]Populate!R2C16:R41C18,3,FALSE)"
            Else
            TempStr2 = "=VLOOKUP(" & distnum & ",[populate.xlsm]Populate!R2C16:R41C18,3,FALSE)"
            End If
            Range("P5").FormulaR1C1 = TempStr2
            Range("P5").Select
            Selection.Copy
            Range("P5").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
              :=False, Transpose:=False
            Range("P5") = Range("P5") & " District"
            End If
         'concatenate county and district
            If Range("P5") <> "" Then
            Range("O4") = Range("O4") & " , " & Range("P5")
            Range("P5") = ""
            Else
            Range("O4") = Range("O4")
            End If
         'bringing over review number
            Range("A10") = Left(Sheets("Temp").Range("A" & i), 1)
            Range("B10") = Mid(Sheets("Temp").Range("A" & i), 2, 1)
            Range("C10") = Mid(Sheets("Temp").Range("A" & i), 3, 1)
            Range("D10") = Mid(Sheets("Temp").Range("A" & i), 4, 1)
            Range("E10") = Mid(Sheets("Temp").Range("A" & i), 5, 1)
            Range("F10") = Right(Sheets("Temp").Range("A" & i), 1)
         'move review number to other pages
          Range("A10:F10").Select
          Selection.Copy
          Range("AK45").Select
           Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
          Range("A10:F10").Select
          Selection.Copy
          Range("AK89").Select
          Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
          Range("A10:F10").Select
          Selection.Copy
          Range("AC132").Select
          Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
         'bringing over case number
         TempStr = Sheets("Temp").Range("F" & i)
         strlen = Len(TempStr)
         If strlen < 7 Then
            For j = 1 To 7 - strlen
                TempStr = "0" & TempStr
            Next j
         End If
            Range("H10") = TempStr
         'bringing over case category
         TempStr = Sheets("Temp").Range("G" & i)
         strlen = Len(TempStr)
         If strlen < 3 Then
            For j = 1 To 3 - strlen
                TempStr = " " & TempStr
            Next j
         End If
            Range("Q10") = TempStr
        'bringing over first and last name
             'Range("B2") = Sheets("Temp").Range("I" & i) & " " & Sheets("Temp").Range("H" & i)
             Range("B2") = Sheets("Temp").Range("H" & i)
         'bringing over address
             Range("B3") = Sheets("Temp").Range("J" & i)
         'bringing over city, state, and zip
             Range("B5") = Sheets("Temp").Range("K" & i) & ",  " & Sheets("Temp").Range("L" & i) _
             & " " & Sheets("Temp").Range("M" & i)
         'bringing over second address
             Range("B4") = ""
             If Sheets("Temp").Range("Z" & i) <> "0" Then
             Range("B4") = Sheets("Temp").Range("Z" & i)
             End If
        'bringing over Program Status Code
             Range("V10") = "00"
               If Sheets("Temp").Range("X" & i) <> "0" Then
                  Range("V10") = Sheets("Temp").Range("X" & i)
               End If
             'Range("W10") = "0"
               'If Sheets("Temp").Range("X" & i) <> "0" Then
                  'Range("W10") = Right(Sheets("Temp").Range("X" & i), 1)
               'End If
        'bringing over Grant Group
             Range("Z10") = ""
               If Sheets("Temp").Range("Y" & i) <> "0" Then
                  Range("Z10") = Sheets("Temp").Range("Y" & i)
               End If
             Range("A1").Select
         'bringing over Total Amount
             Range("AN16") = ""
               If Sheets("Temp").Range("AG" & i) <> "0" Then
                  Range("AN16") = Sheets("Temp").Range("AG" & i)
               End If
             Range("A1").Select
         Next i

End Sub

Sub Fspos()
         Sheets("FS Pos").Name = Sheets("Temp").Range("A1")

         For i = 1 To maxrow
         strTemp = "Processing Form " & i & " of " & maxrow & ". Please be patient..."
         Application.StatusBar = strTemp
            If i > 1 Then
            TempStr = Sheets("Temp").Range("A" & i - 1)
            Sheets(TempStr).Select
            Sheets(TempStr).Copy before:=Sheets("Temp")
            TempStr = TempStr & " (2)"
            Sheets(TempStr).Name = Sheets("Temp").Range("A" & i)
            End If
            
            ' Save, close and open the schedule workbook after 50 worksheets are added.
         ' This prevents the copy/paste operation from hanging
            If i Mod 50 = 0 Then
            
                Application.DisplayAlerts = False
                Windows(fileName).Activate
                ActiveWorkbook.SAVE
                Windows("populate.xlsm").Activate
                Windows(fileName).Close (False)
                Workbooks.Open fileName, UpdateLinks:=False
                Application.DisplayAlerts = True
                 
            End If
          
          If Range("CA1") = "" Then
            Range("CA1") = Format(Date, "short date")
          End If
            
          'looking up district and bringing over district
          Range("E155") = ""
          Range("F155") = ""
          Select Case Sheets("Temp").Range("D" & i)
            Case 2
                TempStr2 = _
                "=LEFT(VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R2C15:R9C16,2,FALSE),1)"
                Range("E155").FormulaR1C1 = TempStr2
                tempstr3 = _
                "=RIGHT(VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R2C15:R9C16,2,FALSE),1)"
                Range("F155").FormulaR1C1 = tempstr3
  '           Case 9
  '              TempStr2 = _
  '              "=LEFT(VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R10C15:R11C16,2,FALSE),1)"
  '              Range("B153").FormulaR1C1 = TempStr2
  '              tempstr3 = _
  '              "=RIGHT(VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R10C15:R11C16,2,FALSE),1)"
  '              Range("C153").FormulaR1C1 = tempstr3
            Case 23
                TempStr2 = _
                "=LEFT(VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R11C15:R13C16,2,FALSE),1)"
                Range("E155").FormulaR1C1 = TempStr2
                tempstr3 = _
                "=RIGHT(VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R11C15:R13C16,2,FALSE),1)"
                Range("F155").FormulaR1C1 = tempstr3
            Case 40
                TempStr2 = _
                "=LEFT(VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R14C15:R15C16,2,FALSE),1)"
                Range("E155").FormulaR1C1 = TempStr2
                tempstr3 = _
                "=RIGHT(VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R14C15:R15C16,2,FALSE),1)"
                Range("F155").FormulaR1C1 = tempstr3
 '           Case 46
 '               TempStr2 = _
 '               "=LEFT(VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R16C15:R17C16,2,FALSE),1)"
 '               Range("B153").FormulaR1C1 = TempStr2
 '               tempstr3 = _
 '               "=RIGHT(VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R16C15:R17C16,2,FALSE),1)"
 '               Range("C153").FormulaR1C1 = tempstr3
            Case 51
                TempStr2 = _
                "=LEFT(VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R18C15:R36C16,2,FALSE),1)"
                Range("E155").FormulaR1C1 = TempStr2
                tempstr3 = _
                "=RIGHT(VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R18C15:R36C16,2,FALSE),1)"
                Range("F155").FormulaR1C1 = tempstr3
            Case 63
                TempStr2 = _
                "=LEFT(VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R37C15:R38C16,2,FALSE),1)"
                Range("E155").FormulaR1C1 = TempStr2
                tempstr3 = _
                "=RIGHT(VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R37C15:R38C16,2,FALSE),1)"
                Range("F155").FormulaR1C1 = tempstr3
'            Case 65
'                TempStr2 = _
'                "=LEFT(VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R39C15:R41C16,2,FALSE),1)"
'                Range("B153").FormulaR1C1 = TempStr2
'                tempstr3 = _
'                "=RIGHT(VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R39C15:R41C16,2,FALSE),1)"
'                Range("C153").FormulaR1C1 = tempstr3
          End Select
          Range("E155:F155").Select
          Selection.Copy
          Range("E155:F155").Select
          Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
          If i = 1 Then
            'bringing over month
            monthstr1 = Application.WorksheetFunction.Text(revmonth, "MMYYYY")
            Range("AD18") = Left(monthstr1, 2)
            Range("AG18") = Right(monthstr1, 4)
          End If
          
          'putting a stratum code in
'          If revmonth > 39844 And revmonth < 40058 Then
'            Range("AL18") = "02"
'          End If
          
         'determining if county has only one number or two numbers in spreadsheet and adding district
            countyunit = Sheets("Temp").Range("D" & i)
            If Sheets("Temp").Range("D" & i) < 10 Then
            countyunit = "0" & Sheets("Temp").Range("D" & i)
            End If
        '    Range("X18") = Sheets("Temp").Range("C" & i) & CountyUnit Valerie
            Range("B155") = Left(Sheets("Temp").Range("C" & i), 1)
            Range("C155") = Left(countyunit, 1)
            Range("D155") = Right(Sheets("Temp").Range("D" & i), 1)
         'putting county name
            TempStr = "=VLOOKUP(Temp!R" & i & "C4,[populate.xlsm]Populate!R2C30:R68C31,2,FALSE)"
            Range("M5").FormulaR1C1 = TempStr
            Range("M5").Select
            Selection.Copy
            Range("M5").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
              :=False, Transpose:=False
            Range("M5") = Range("M5") & " County"
         'district name
            Range("O7") = ""
            distnum = Range("E155") & Range("F155")
            If distnum <> "" Then
            If Val(distnum) < 11 Then
            TempStr2 = "=VLOOKUP(" & """" & distnum & """" & ",[populate.xlsm]Populate!R2C16:R41C18,3,FALSE)"
            Else
            TempStr2 = "=VLOOKUP(" & distnum & ",[populate.xlsm]Populate!R2C16:R41C18,3,FALSE)"
            End If
            Range("O7").FormulaR1C1 = TempStr2
            Range("O7").Select
            Selection.Copy
            Range("O7").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
              :=False, Transpose:=False
            Range("O7") = Range("O7") & " District"
            End If
         'concatenate county and district name
          If Range("O7") <> "" Then
            Range("M5") = Range("M5") & " , " & Range("O7")
            Range("O7") = ""
            Else
            Range("M5") = Range("M5")
            End If
         'concatenate county and district
            countyunit = Sheets("Temp").Range("D" & i)
            If Sheets("Temp").Range("D" & i) < 10 Then
            countyunit = "0" & Sheets("Temp").Range("D" & i)
            End If
         'bringing over review number
            TempStr = Sheets("Temp").Range("A" & i)
            strlen = Len(TempStr)
            If strlen < 6 Then
                For j = 1 To 6 - strlen
                    TempStr = "0" & TempStr
                Next j
            End If
             Range("A18") = TempStr
         'bringing over case number
            TempStr = Sheets("Temp").Range("F" & i)
            strlen = Len(TempStr)
         If strlen < 7 Then
            For j = 1 To 7 - strlen
                TempStr = "0" & TempStr
            Next j
         End If
         If Left(Sheets("Temp").Range("A" & i), 2) = "65" Or _
            Left(Sheets("Temp").Range("A" & i), 2) = "66" Then
            TempStr = "00A" & TempStr
         Else
            TempStr = "00" & TempStr & Right(Sheets("Temp").Range("G" & i), 1)
         End If
             Range("I18") = TempStr
         'bringing over first and last name
             Range("B4") = Sheets("Temp").Range("I" & i) & " " & Sheets("Temp").Range("H" & i)
         'bringing over address
             TempStr = Sheets("Temp").Range("J" & i)
             If Sheets("Temp").Range("Z" & i) <> "0" Then
             TempStr = TempStr
             End If
             Range("B5") = TempStr
         'bringing over city, state, and zip
             Range("B7") = Sheets("Temp").Range("K" & i) & ",  " & Sheets("Temp").Range("L" & i) _
             & " " & Sheets("Temp").Range("M" & i)
         'bringing over second address
             Range("B6") = ""
             If Sheets("Temp").Range("Z" & i) <> "0" Then
             Range("B6") = Sheets("Temp").Range("Z" & i)
             End If
        'putting name in examiner field.  Looking up examiner number
          TempStr4 = _
          "=VLOOKUP(Temp!R" & i & "C14,[populate.xlsm]Populate!R2C27:R39C28,2,FALSE)"
                Range("AE5").FormulaR1C1 = TempStr4
          Range("AE5").Select
          Selection.Copy
          Range("AE5").Select
          Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
         'bringing over examiner number
            If (Len(Sheets("Temp").Range("N" & i))) < 2 Then
                Range("AJ5") = "0"
            Else
             Range("AJ5") = Left(Sheets("Temp").Range("N" & i), 1)
            End If
            Range("AK5") = Right(Sheets("Temp").Range("N" & i), 1)
         'putting county name
            TempStr = "=VLOOKUP(Temp!R" & i & "C4,[populate.xlsm]Populate!R2C30:R68C31,2,FALSE)"
            Range("P5").FormulaR1C1 = TempStr
            Range("P5").Select
            Selection.Copy
            Range("P5").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
              :=False, Transpose:=False
            Range("P5") = Range("P5") & " County"
            Range("A1").Select
        Next i
End Sub
Sub Fsneg()
         Sheets("FS Neg").Name = Sheets("Temp").Range("A1")
         For i = 1 To maxrow
         strTemp = "Processing Form " & i & " of " & maxrow & ". Please be patient..."
         Application.StatusBar = strTemp
            If i > 1 Then
            TempStr = Sheets("Temp").Range("A" & i - 1)
            Sheets(TempStr).Select
            Sheets(TempStr).Copy before:=Sheets("Temp")
            TempStr = TempStr & " (2)"
            Sheets(TempStr).Name = Sheets("Temp").Range("A" & i)
            End If
            
             ' Save, close and open the schedule workbook after 50 worksheets are added.
         ' This prevents the copy/paste operation from hanging
            If i Mod 50 = 0 Then
            
                Application.DisplayAlerts = False
                Windows(fileName).Activate
                ActiveWorkbook.SAVE
                Windows("populate.xlsm").Activate
                Windows(fileName).Close (False)
                Workbooks.Open fileName, UpdateLinks:=False
                Application.DisplayAlerts = True
                 
            End If
            
          If Range("AQ1") = "" Then
            Range("AQ1") = Format(Date, "short date")
          End If
          
          If i = 1 Then
            'bringing over month
            monthstr1 = Application.WorksheetFunction.Text(revmonth, "MMYYYY")
            Range("AF20") = Left(monthstr1, 2)
            Range("AI20") = Right(monthstr1, 4)
          End If
          
          'putting a stratum code in
          'If revmonth > 39844 And revmonth < 40058 Then
          '  Range("C24") = "02"
          'End If
          
         'determining if county has only one number or two numbers in spreadsheet
           countyunit = Sheets("Temp").Range("D" & i)
            If Sheets("Temp").Range("D" & i) < 10 Then
            countyunit = "0" & Sheets("Temp").Range("D" & i)
            End If
           ' Range("AA20") = Sheets("Temp").Range("C" & i) & CountyUnit
             Range("E56") = Sheets("Temp").Range("C" & i)
             Range("F56") = Left(countyunit, 1)
             Range("G56") = Right(countyunit, 1)
             
          'looking up district and bringing over district
          Range("AD20") = ""
          Range("AE20") = ""
          Select Case Sheets("Temp").Range("D" & i)
            Case 2
                TempStr2 = _
                "=LEFT(VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R2C15:R9C16,2,FALSE),1)"
                Range("AD20").FormulaR1C1 = TempStr2
                tempstr3 = _
                "=RIGHT(VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R2C15:R9C16,2,FALSE),1)"
                Range("AE20").FormulaR1C1 = tempstr3
 '            Case 9
 '               TempStr2 = _
 '               "=LEFT(VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R10C15:R11C16,2,FALSE),1)"
 '               Range("AD20").FormulaR1C1 = TempStr2
 '               tempstr3 = _
 '               "=RIGHT(VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R10C15:R11C16,2,FALSE),1)"
 '               Range("AE20").FormulaR1C1 = tempstr3
            Case 23
                TempStr2 = _
                "=LEFT(VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R11C15:R13C16,2,FALSE),1)"
                Range("AD20").FormulaR1C1 = TempStr2
                tempstr3 = _
                "=RIGHT(VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R11C15:R13C16,2,FALSE),1)"
                Range("AE20").FormulaR1C1 = tempstr3
            Case 40
                TempStr2 = _
                "=LEFT(VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R14C15:R15C16,2,FALSE),1)"
                Range("AD20").FormulaR1C1 = TempStr2
                tempstr3 = _
                "=RIGHT(VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R14C15:R15C16,2,FALSE),1)"
                Range("AE20").FormulaR1C1 = tempstr3
 '           Case 46
 '               TempStr2 = _
 '               "=LEFT(VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R16C15:R17C16,2,FALSE),1)"
 '               Range("AD20").FormulaR1C1 = TempStr2
 '               tempstr3 = _
 '               "=RIGHT(VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R16C15:R17C16,2,FALSE),1)"
 '               Range("AE20").FormulaR1C1 = tempstr3
            Case 51
                TempStr2 = _
                "=LEFT(VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R18C15:R36C16,2,FALSE),1)"
                Range("AD20").FormulaR1C1 = TempStr2
                tempstr3 = _
                "=RIGHT(VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R18C15:R36C16,2,FALSE),1)"
                Range("AE20").FormulaR1C1 = tempstr3
            Case 63
                TempStr2 = _
                "=LEFT(VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R37C15:R38C16,2,FALSE),1)"
                Range("AD20").FormulaR1C1 = TempStr2
                tempstr3 = _
                "=RIGHT(VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R37C15:R38C16,2,FALSE),1)"
                Range("AE20").FormulaR1C1 = tempstr3
            Case 65
                TempStr2 = _
                "=LEFT(VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R39C15:R41C16,2,FALSE),1)"
                Range("AD20").FormulaR1C1 = TempStr2
                tempstr3 = _
                "=RIGHT(VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R39C15:R41C16,2,FALSE),1)"
                Range("AE20").FormulaR1C1 = tempstr3
          End Select
          Range("AD20:AE20").Select
          Selection.Copy
          Range("AD20:AE20").Select
          Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
            
            'bringing over review number
            TempStr = Sheets("Temp").Range("A" & i)
            strlen = Len(TempStr)
            If strlen < 6 Then
                For j = 1 To 6 - strlen
                    TempStr = "0" & TempStr
                Next j
            End If
             Range("C20") = TempStr
         'bringing over case number
            TempStr = Sheets("Temp").Range("F" & i)
            strlen = Len(TempStr)
         If strlen < 7 Then
            For j = 1 To 7 - strlen
                TempStr = "0" & TempStr
            Next j
         End If
         If Left(Sheets("Temp").Range("A" & i), 2) = "65" Or _
            Left(Sheets("Temp").Range("A" & i), 2) = "66" Then
            TempStr = "00A" & TempStr
         Else
            TempStr = "00" & TempStr & Right(Sheets("Temp").Range("G" & i), 1)
         End If
             Range("L20") = TempStr
         'bringing over first and last name
             Range("C8") = Sheets("Temp").Range("I" & i) & " " & Sheets("Temp").Range("H" & i)
         'bringing over address
             TempStr = Sheets("Temp").Range("J" & i)
             If Sheets("Temp").Range("Z" & i) <> "0" Then
             TempStr = TempStr & " , " & Sheets("Temp").Range("Z" & i)
             End If
             Range("C12") = TempStr
         'bringing over city, state, and zip
             Range("C13") = Sheets("Temp").Range("K" & i) & ",  " & Sheets("Temp").Range("L" & i) _
             & " " & Sheets("Temp").Range("M" & i)
         'putting name in examiner field.  Looking up examiner number
          TempStr4 = _
          "=VLOOKUP(Temp!R" & i & "C14,[populate.xlsm]Populate!R2C27:R39C28,2,FALSE)"
                Range("Q17").FormulaR1C1 = TempStr4
          Range("Q17").Select
          Selection.Copy
          Range("Q17").Select
          Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
         'bringing over examiner number
            If (Len(Sheets("Temp").Range("N" & i))) < 2 Then
                Range("W17") = "0" & Sheets("Temp").Range("N" & i)
            Else
                Range("W17") = Sheets("Temp").Range("N" & i)
            End If
        'putting county name
            TempStr = "=VLOOKUP(Temp!R" & i & "C4,[populate.xlsm]Populate!R2C30:R68C31,2,FALSE)"
            Range("c1").FormulaR1C1 = TempStr
            Range("c1").Select
            Selection.Copy
            Range("c1").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
              :=False, Transpose:=False
            Range("c1") = Range("c1") & " County"
        Select Case Sheets("Temp").Range("D" & i)
            Case 2
                TempStr2 = _
                "=VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R2C15:R9C18,4,FALSE)"
                Range("c6").FormulaR1C1 = TempStr2
  '           Case 9
  '              TempStr2 = _
  '              "=VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R10C15:R11C18,4,FALSE)"
  '              Range("C6").FormulaR1C1 = TempStr2
            Case 23
                TempStr2 = _
                "=VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R11C15:R13C18,4,FALSE)"
                Range("C6").FormulaR1C1 = TempStr2
            Case 40
                TempStr2 = _
                "=VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R14C15:R15C18,4,FALSE)"
                Range("C6").FormulaR1C1 = TempStr2
  '          Case 46
  '              TempStr2 = _
  '              "=VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R16C15:R17C18,4,FALSE)"
  '              Range("C6").FormulaR1C1 = TempStr2
            Case 51
                TempStr2 = _
                "=VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R18C15:R36C18,4,FALSE)"
                Range("C6").FormulaR1C1 = TempStr2
            Case 63
                TempStr2 = _
                "=VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R37C15:R38C18,4,FALSE)"
                Range("C6").FormulaR1C1 = TempStr2
            Case 65
                TempStr2 = _
                "=VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R39C15:R41C18,4,FALSE)"
                Range("C6").FormulaR1C1 = TempStr2
          End Select
          Range("C6").Select
          Selection.Copy
          Range("C6").Select
          Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
             :=False, Transpose:=False
         'concatenate county and district
            If Range("c6") <> "" Then
            Range("c6") = Range("c6") & " District"
            Range("c1") = Range("c1") & " , " & Range("c6")
            Range("c6") = ""
            Else
            Range("c1") = Range("c1")
            End If
          'putting whether Rejected, closed, or suspended
          'If Left(Range("c20"), 2) = "65" Or Left(Range("c20"), 2) = "66" Then
          '      Range("AJ4") = "Rejected Application"
          '      Range("AE24") = 1
          'End If
          'If Left(Range("c20"), 3) = "060" Or Left(Range("c20"), 3) = "061" Then
          '      If Sheets("Temp").Range("AC" & i) = "C" Then
          '          Range("AJ4") = "Closed Application"
          '          Range("AE24") = 2
          '      ElseIf Sheets("Temp").Range("AC" & i) = "S" Then
          '          Range("AJ4") = "Suspended Appl."
          '          Range("AE24") = 3
          '      End If
          'End If
          
         Range("A1").Select
         Next i
End Sub
Sub tanf()
         If program = "TANF" Then
           Sheets("TANF").Name = Sheets("Temp").Range("A1")
         Else
           Sheets("TANF CAR").Name = Sheets("Temp").Range("A1")
         End If
         For i = 1 To maxrow
         strTemp = "Processing Form " & i & " of " & maxrow & ". Please be patient..."
         Application.StatusBar = strTemp
         
         ' Save, close and open the schedule workbook after 50 worksheets are added.
         ' This prevents the copy/paste operation from hanging
            If i Mod 50 = 0 Then
            
                Application.DisplayAlerts = False
                Windows(fileName).Activate
                ActiveWorkbook.SAVE
                Windows("populate.xlsm").Activate
                Windows(fileName).Close (False)
                Workbooks.Open fileName, UpdateLinks:=False
                Application.DisplayAlerts = True
                 
            End If
                
           If Range("AU1") = "" Then
            Range("AU1") = Format(Date, "short date")
          End If
                
            If i > 1 Then
            TempStr = Sheets("Temp").Range("A" & i - 1)
            Sheets(TempStr).Select
            Sheets(TempStr).Copy After:=Sheets(TempStr)
            TempStr = TempStr & " (2)"
            Sheets(TempStr).Name = Sheets("Temp").Range("A" & i)
            End If
                'looking up district and bringing over district
          Range("Y10") = ""
          Select Case Sheets("Temp").Range("D" & i)
            Case 2
                TempStr2 = _
                "=VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R2C15:R9C16,2,FALSE)"
                Range("Y10").FormulaR1C1 = TempStr2
 '            Case 9
 '              TempStr2 = _
 '               "=VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R10C15:R11C16,2,FALSE)"
 '               Range("Y10").FormulaR1C1 = TempStr2
            Case 23
                TempStr2 = _
                "=VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R11C15:R13C16,2,FALSE)"
                Range("Y10").FormulaR1C1 = TempStr2
            Case 40
               TempStr2 = _
                "=VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R14C15:R15C16,2,FALSE)"
                Range("Y10").FormulaR1C1 = TempStr2
 '           Case 46
 '              TempStr2 = _
 '               "=VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R16C15:R17C16,2,FALSE)"
 '               Range("Y10").FormulaR1C1 = TempStr2
            Case 51
               TempStr2 = _
                "=VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R18C15:R36C16,2,FALSE)"
                Range("Y10").FormulaR1C1 = TempStr2
            Case 63
               TempStr2 = _
                "=VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R37C15:R38C16,2,FALSE)"
                Range("Y10").FormulaR1C1 = TempStr2
            Case 65
               TempStr2 = _
                "=VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R39C15:R41C16,2,FALSE)"
                Range("Y10").FormulaR1C1 = TempStr2
          End Select
          Range("Y10").Select
          Selection.Copy
          Range("Y10").Select
          Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
          'putting name in examiner field.  Looking up examiner number
          TempStr4 = _
          "=VLOOKUP(Temp!R" & i & "C14,[populate.xlsm]Populate!R2C27:R39C28,2,FALSE)"
                Range("AG2").FormulaR1C1 = TempStr4
          Range("AG2").Select
          Selection.Copy
          Range("AG2").Select
          Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
          If i = 1 Then
            'bringing over month
            monthstr1 = Application.WorksheetFunction.Text(revmonth, "MMYYYY")
            Range("AB10") = monthstr1
          End If
         'bringing over examiner number
            If (Len(Sheets("Temp").Range("N" & i))) < 2 Then
                Range("AO3") = "0"
            Else
             Range("AO3") = Left(Sheets("Temp").Range("N" & i), 1)
            End If
             Range("AP3") = Right(Sheets("Temp").Range("N" & i), 1)
         'determining if county has only one number or two numbers in spreadsheet
          countyunit = 0
            If Sheets("Temp").Range("D" & i) > 9 Then
            countyunit = Sheets("Temp").Range("D" & i)
            Else
            countyunit = countyunit & Sheets("Temp").Range("D" & i)
            End If
            Range("U10") = Sheets("Temp").Range("C" & i) & countyunit
         'putting county name
            TempStr = "=VLOOKUP(Temp!R" & i & "C4,[populate.xlsm]Populate!R2C30:R68C31,2,FALSE)"
            Range("O4").FormulaR1C1 = TempStr
            Range("O4").Select
            Selection.Copy
            Range("O4").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
              :=False, Transpose:=False
            Range("O4") = Range("O4") & " County"
        'district name
            Range("P5") = ""
            distnum = Range("Y10") & Range("Z10")
            If distnum <> "" Then
            If Val(distnum) < 11 Then
            TempStr2 = "=VLOOKUP(" & """" & distnum & """" & ",[populate.xlsm]Populate!R2C16:R41C18,3,FALSE)"
            Else
            TempStr2 = "=VLOOKUP(" & distnum & ",[populate.xlsm]Populate!R2C16:R41C18,3,FALSE)"
            End If
            Range("P5").FormulaR1C1 = TempStr2
            Range("P5").Select
            Selection.Copy
            Range("P5").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
              :=False, Transpose:=False
            Range("P5") = Range("P5") & " District"
            End If
         'concatenate county and district
            If Range("P5") <> "" Then
            Range("O4") = Range("O4") & " , " & Range("P5")
            Range("P5") = ""
            Else
            Range("O4") = Range("O4")
            End If
         'bringing over review number
            Range("A10") = Sheets("Temp").Range("A" & i)
         'bringing over case number
         TempStr = Sheets("Temp").Range("F" & i)
         strlen = Len(TempStr)
         If strlen < 7 Then
            For j = 1 To 7 - strlen
                TempStr = "0" & TempStr
            Next j
         End If
            Range("I10") = TempStr
         'bringing over case category
             Range("Q10") = Sheets("Temp").Range("G" & i)
         'bringing over first and last name
             Range("B2") = Sheets("Temp").Range("I" & i) & " " & Sheets("Temp").Range("H" & i)
         'bringing over Grant Group
             Range("S10") = ""
               If Sheets("Temp").Range("X" & i) <> "0" Then
                  Range("S10") = Sheets("Temp").Range("X" & i)
               End If
         'bringing over address
             Range("B3") = Sheets("Temp").Range("J" & i)
         'bringing over city, state, and zip
             Range("B4") = Sheets("Temp").Range("K" & i) & ",  " & Sheets("Temp").Range("L" & i) _
             & " " & Sheets("Temp").Range("M" & i)
         'Bringing over food stamp allotment
             Range("B24") = Sheets("Temp").Range("AC" & i)
         'Clearing out old line numbers
            jrow = 30
            Do Until (Cells(jrow, 1) = "")
                Cells(jrow, 1) = ""
                Cells(jrow, 13) = ""
                Cells(jrow, 16) = ""
                jrow = jrow + 2
            Loop
         'Bringing over line number
            jrow = 30
            jcol = 30
            'If Sheets("Temp").Cells(i, jcol) = "" Then jrow = ""
            Do Until (Sheets("Temp").Cells(i, jcol) = "" Or jrow > 44)
            If Sheets("Temp").Cells(i, jcol) < 10 Then
                Cells(jrow, 1) = "0" & Sheets("Temp").Cells(i, jcol)
            Else
                 Cells(jrow, 1) = Sheets("Temp").Cells(i, jcol)
            End If
         'Bringing over Male or Female
                Cells(jrow, 16) = 2
                If Sheets("Temp").Cells(i, jcol + 1) = "M" Then Cells(jrow, 16) = 1
         'Calculating age
                Cells(jrow, 13) = Int((Now() - Sheets("Temp").Cells(i, jcol + 4)) / 365)
                jrow = jrow + 2
                jcol = jcol + 6
            Loop
         Range("A1").Select
         Next i
End Sub
Sub Maneg()
        Sheets("MA Neg").Name = Sheets("Temp").Range("A1")
         For i = 1 To maxrow
            If i > 1 Then
            TempStr = Sheets("Temp").Range("A" & i - 1)
            Sheets(TempStr).Select
            Sheets(TempStr).Copy before:=Sheets("Temp")
            TempStr = TempStr & " (2)"
            Sheets(TempStr).Name = Sheets("Temp").Range("A" & i)
            End If
            
            'bringing over month - just put month on first sheet, then it is copied
          If i = 1 Then
            monthstr1 = Application.WorksheetFunction.Text(revmonth, "MMYYYY")
            Range("T11") = monthstr1
          End If
          'If i = 1 Then
            'bringing over month
            'Range("T11") = monthstr
          'End If
          
          If Range("AI1") = "" Then
            Range("AI1") = Format(Date, "short date")
          End If
           
           If Range("B11") = "" Then
            Range("B11") = Format(Date, "short date")
          End If
          
         'bringing over examiner number
          If (Len(Sheets("Temp").Range("N" & i))) < 2 Then
                Cells(11, 28) = "0"
            Else
             Cells(11, 28) = Left(Sheets("Temp").Range("N" & i), 1)
            End If
            Cells(11, 29) = Right(Sheets("Temp").Range("N" & i), 1)
         'determining if county has only one number or two numbers in spreadsheet
            countyunit = 0
            If Sheets("Temp").Range("D" & i) > 9 Then
            countyunit = Left(Sheets("Temp").Range("D" & i), 1)
            End If
            Range("G15") = Sheets("Temp").Range("C" & i) & countyunit & Right(Sheets("Temp").Range("D" & i), 1)

         'putting county name
            TempStr = "=VLOOKUP(Temp!R" & i & "C4,[populate.xlsm]Populate!R2C30:R68C31,2,FALSE)"
            Range("F2").FormulaR1C1 = TempStr
            Range("F2").Select
            Selection.Copy
            Range("F2").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
              :=False, Transpose:=False
            Range("F2") = Range("F2") & " County"
            
            'looking up district and bringing over district
          Range("J15") = ""
          Range("K15") = ""
          Select Case Sheets("Temp").Range("D" & i)
            Case 2
                TempStr2 = _
                "=LEFT(VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R2C15:R9C16,2,FALSE),1)"
                Range("J15").FormulaR1C1 = TempStr2
                tempstr3 = _
                "=RIGHT(VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R2C15:R9C16,2,FALSE),1)"
                Range("K15").FormulaR1C1 = tempstr3
   '          Case 9
   '             TempStr2 = _
   '             "=LEFT(VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R10C15:R11C16,2,FALSE),1)"
   '             Range("J15").FormulaR1C1 = TempStr2
   '             tempstr3 = _
   '             "=RIGHT(VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R10C15:R11C16,2,FALSE),1)"
   '             Range("K15").FormulaR1C1 = tempstr3
            Case 23
                TempStr2 = _
                "=LEFT(VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R11C15:R13C16,2,FALSE),1)"
                Range("J15").FormulaR1C1 = TempStr2
                tempstr3 = _
                "=RIGHT(VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R11C15:R13C16,2,FALSE),1)"
                Range("K15").FormulaR1C1 = tempstr3
            Case 40
                TempStr2 = _
                "=LEFT(VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R14C15:R15C16,2,FALSE),1)"
                Range("J15").FormulaR1C1 = TempStr2
                tempstr3 = _
                "=RIGHT(VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R14C15:R15C16,2,FALSE),1)"
                Range("K15").FormulaR1C1 = tempstr3
   '         Case 46
   '             TempStr2 = _
   '             "=LEFT(VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R16C15:R17C16,2,FALSE),1)"
   '             Range("J15").FormulaR1C1 = TempStr2
   '             tempstr3 = _
   '             "=RIGHT(VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R16C15:R17C16,2,FALSE),1)"
   '             Range("K15").FormulaR1C1 = tempstr3
            Case 51
                TempStr2 = _
                "=LEFT(VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R18C15:R36C16,2,FALSE),1)"
                Range("J15").FormulaR1C1 = TempStr2
                tempstr3 = _
                "=RIGHT(VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R18C15:R36C16,2,FALSE),1)"
                Range("K15").FormulaR1C1 = tempstr3
            Case 63
                TempStr2 = _
                "=LEFT(VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R37C15:R38C16,2,FALSE),1)"
                Range("J15").FormulaR1C1 = TempStr2
                tempstr3 = _
                "=RIGHT(VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R37C15:R38C16,2,FALSE),1)"
                Range("K15").FormulaR1C1 = tempstr3
            Case 65
                TempStr2 = _
                "=LEFT(VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R39C15:R41C16,2,FALSE),1)"
                Range("J15").FormulaR1C1 = TempStr2
                tempstr3 = _
                "=RIGHT(VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R39C15:R41C16,2,FALSE),1)"
                Range("K15").FormulaR1C1 = tempstr3
          End Select
          Range("J15:K15").Select
          Selection.Copy
          Range("J15:K15").Select
          Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
            
            'district name
            Range("W2") = ""
            distnum = Range("J15") & Range("K15")
            If distnum <> "" Then
            If Val(distnum) < 11 Then
            TempStr2 = "=VLOOKUP(" & """" & distnum & """" & ",[populate.xlsm]Populate!R2C16:R41C18,3,FALSE)"
            Else
            TempStr2 = "=VLOOKUP(" & distnum & ",[populate.xlsm]Populate!R2C16:R41C18,3,FALSE)"
            End If
            Range("W2").FormulaR1C1 = TempStr2
            Range("W2").Select
            Selection.Copy
            Range("W2").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
              :=False, Transpose:=False
            Range("W2") = Range("W2") & " District"
            End If
         'concatenate county and district
            If Range("W2") <> "" Then
            Range("F2") = Range("F2") & " , " & Range("W2")
            Range("W2") = ""
            Else
            Range("F2") = Range("F2")
            End If
            

            
        'bringing over review number
            Range("L15") = Sheets("Temp").Range("A" & i)
         'bringing over case number
            TempStr = Sheets("Temp").Range("F" & i)
            strlen = Len(TempStr)
         If strlen < 7 Then
             For j = 1 To 7 - strlen
              If Left(Sheets("Temp").Range("A" & i), 2) = "81" Or _
              Left(Sheets("Temp").Range("A" & i), 2) = "83" Then
                TempStr = "A0" & TempStr
              Else
                TempStr = "0" & TempStr
              End If
             Next j
         End If
         Range("S15") = TempStr
         'bringing over case category
         TempStr = Sheets("Temp").Range("G" & i)
         strlen = Len(TempStr)
         If strlen < 3 Then
            For j = 1 To 3 - strlen
                TempStr = " " & TempStr
            Next j
         End If
             Range("AB15") = TempStr
         'bringing over first and last name
             Range("B7") = Sheets("Temp").Range("I" & i) & " " & Sheets("Temp").Range("H" & i)
         'bringing over address
             Range("L6") = Sheets("Temp").Range("J" & i)
         'bringing over city, state, and zip
             Range("L8") = Sheets("Temp").Range("K" & i) & ",  " & Sheets("Temp").Range("L" & i) _
             & " " & Sheets("Temp").Range("M" & i)
         'bringing over Program Status Code
             Range("B19") = Sheets("Temp").Range("X" & i)
         'bringing over Grant Group
             Range("F19") = Sheets("Temp").Range("Y" & i)
         'bringing over second address
             Range("L7") = ""
             If Sheets("Temp").Range("Z" & i) <> "0" Then
             Range("L7") = Sheets("Temp").Range("Z" & i)
             End If
         'determining if case is rejected or closed
              If Left(Sheets("Temp").Range("A" & i), 2) = "81" Or _
                Left(Sheets("Temp").Range("A" & i), 2) = "83" Then
                Range("AF19") = 1
              Else
                Range("AF19") = 2
              End If
            'If Left(Sheets("Temp").Range("A" & i), 2) = "81" Or _
            '    Left(Sheets("Temp").Range("A" & i), 2) = "83" Then
            '    Range("C25") = 5
            '  End If
             Range("A1").Select
         Next i
End Sub
Sub PE()
        Sheets("PE Review").Name = Sheets("Temp").Range("A1")
         For i = 1 To maxrow
            If i > 1 Then
            TempStr = Sheets("Temp").Range("A" & i - 1)
            Sheets(TempStr).Select
            Sheets(TempStr).Copy before:=Sheets("Temp")
            TempStr = TempStr & " (2)"
            Sheets(TempStr).Name = Sheets("Temp").Range("A" & i)
            End If
            
            'bringing over month - just put month on first sheet, then it is copied
          If i = 1 Then
            monthstr1 = Application.WorksheetFunction.Text(revmonth, "MMYYYY")
            Range("T11") = monthstr1
          End If
          'If i = 1 Then
            'bringing over month
            'Range("T11") = monthstr
          'End If
          
         'bringing over examiner number
          If (Len(Sheets("Temp").Range("N" & i))) < 2 Then
                Cells(11, 28) = "0"
            Else
             Cells(11, 28) = Left(Sheets("Temp").Range("N" & i), 1)
            End If
            Cells(11, 29) = Right(Sheets("Temp").Range("N" & i), 1)
         'determining if county has only one number or two numbers in spreadsheet
            countyunit = 0
            If Sheets("Temp").Range("D" & i) > 9 Then
            countyunit = Left(Sheets("Temp").Range("D" & i), 1)
            End If
            Range("G15") = Sheets("Temp").Range("C" & i) & countyunit & Right(Sheets("Temp").Range("D" & i), 1)

         'putting county name
            TempStr = "=VLOOKUP(Temp!R" & i & "C4,[populate.xlsm]Populate!R2C30:R68C31,2,FALSE)"
            Range("F2").FormulaR1C1 = TempStr
            Range("F2").Select
            Selection.Copy
            Range("F2").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
              :=False, Transpose:=False
            Range("F2") = Range("F2") & " County"
            
            'looking up district and bringing over district
          Range("J15") = ""
          Range("K15") = ""
          Select Case Sheets("Temp").Range("D" & i)
            Case 2
                TempStr2 = _
                "=LEFT(VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R2C15:R9C16,2,FALSE),1)"
                Range("J15").FormulaR1C1 = TempStr2
                tempstr3 = _
                "=RIGHT(VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R2C15:R9C16,2,FALSE),1)"
                Range("K15").FormulaR1C1 = tempstr3
             Case 9
                TempStr2 = _
                "=LEFT(VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R10C15:R11C16,2,FALSE),1)"
                Range("J15").FormulaR1C1 = TempStr2
                tempstr3 = _
                "=RIGHT(VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R10C15:R11C16,2,FALSE),1)"
                Range("K15").FormulaR1C1 = tempstr3
            Case 23
                TempStr2 = _
                "=LEFT(VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R11C15:R13C16,2,FALSE),1)"
                Range("J15").FormulaR1C1 = TempStr2
                tempstr3 = _
                "=RIGHT(VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R11C15:R13C16,2,FALSE),1)"
                Range("K15").FormulaR1C1 = tempstr3
            Case 40
                TempStr2 = _
                "=LEFT(VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R14C15:R15C16,2,FALSE),1)"
                Range("J15").FormulaR1C1 = TempStr2
                tempstr3 = _
                "=RIGHT(VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R14C15:R15C16,2,FALSE),1)"
                Range("K15").FormulaR1C1 = tempstr3
            Case 46
                TempStr2 = _
                "=LEFT(VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R16C15:R17C16,2,FALSE),1)"
                Range("J15").FormulaR1C1 = TempStr2
                tempstr3 = _
                "=RIGHT(VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R16C15:R17C16,2,FALSE),1)"
                Range("K15").FormulaR1C1 = tempstr3
            Case 51
                TempStr2 = _
                "=LEFT(VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R18C15:R36C16,2,FALSE),1)"
                Range("J15").FormulaR1C1 = TempStr2
                tempstr3 = _
                "=RIGHT(VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R18C15:R36C16,2,FALSE),1)"
                Range("K15").FormulaR1C1 = tempstr3
            Case 63
                TempStr2 = _
                "=LEFT(VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R37C15:R38C16,2,FALSE),1)"
                Range("J15").FormulaR1C1 = TempStr2
                tempstr3 = _
                "=RIGHT(VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R37C15:R38C16,2,FALSE),1)"
                Range("K15").FormulaR1C1 = tempstr3
            Case 65
                TempStr2 = _
                "=LEFT(VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R39C15:R41C16,2,FALSE),1)"
                Range("J15").FormulaR1C1 = TempStr2
                tempstr3 = _
                "=RIGHT(VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R39C15:R41C16,2,FALSE),1)"
                Range("K15").FormulaR1C1 = tempstr3
          End Select
          Range("J15:K15").Select
          Selection.Copy
          Range("J15:K15").Select
          Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
            
            'district name
            Range("W2") = ""
            distnum = Range("J15") & Range("K15")
            If distnum <> "" Then
            If Val(distnum) < 11 Then
            TempStr2 = "=VLOOKUP(" & """" & distnum & """" & ",[populate.xlsm]Populate!R2C16:R41C18,3,FALSE)"
            Else
            TempStr2 = "=VLOOKUP(" & distnum & ",[populate.xlsm]Populate!R2C16:R41C18,3,FALSE)"
            End If
            Range("W2").FormulaR1C1 = TempStr2
            Range("W2").Select
            Selection.Copy
            Range("W2").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
              :=False, Transpose:=False
            Range("W2") = Range("W2") & " District"
            End If
         'concatenate county and district
            If Range("W2") <> "" Then
            Range("F2") = Range("F2") & " , " & Range("W2")
            Range("W2") = ""
            Else
            Range("F2") = Range("F2")
            End If
            
        'bringing over review number
            Range("L15") = Sheets("Temp").Range("A" & i)
         'bringing over case number
            TempStr = Sheets("Temp").Range("F" & i)
            strlen = Len(TempStr)
         If strlen < 7 Then
             For j = 1 To 7 - strlen
              If Left(Sheets("Temp").Range("A" & i), 2) = "81" Or _
              Left(Sheets("Temp").Range("A" & i), 2) = "83" Then
                TempStr = "A0" & TempStr
              Else
                TempStr = "0" & TempStr
              End If
             Next j
         End If
         Range("S15") = TempStr
         
         'bringing over case category
         TempStr = Sheets("Temp").Range("G" & i)
         strlen = Len(TempStr)
         If strlen < 3 Then
            For j = 1 To 3 - strlen
                TempStr = " " & TempStr
            Next j
         End If
         'category
             Range("AB15") = TempStr
         'bringing over first and last name
             Range("B7") = Sheets("Temp").Range("I" & i) & " " & Sheets("Temp").Range("H" & i)
         'bringing over address
             Range("L6") = Sheets("Temp").Range("J" & i)
         'bringing over city, state, and zip
             Range("L8") = Sheets("Temp").Range("K" & i) & ",  " & Sheets("Temp").Range("L" & i) _
             & " " & Sheets("Temp").Range("M" & i)
         'bringing over Program Status Code
             Range("B19") = Sheets("Temp").Range("X" & i)
         'bringing over Grant Group
             Range("E19") = Sheets("Temp").Range("Y" & i)
         'bringing over Provider ID
             Range("AB18") = Sheets("Temp").Range("AA" & i)
         'bringing over Provider Name
             Range("AB19") = Sheets("Temp").Range("AB" & i)
         'bringing over Date of Birth
             Range("U19") = Sheets("Temp").Range("AC" & i)
         Next i
End Sub
Public Function EOMONTH(InputDate As Date, Optional MonthsToAdd As Integer)
' Returns the date of the last day of month, a specified number of months
' following a given date.
   Dim TotalMonths As Integer
   Dim NewMonth As Integer
   Dim NewYear As Integer

   If IsMissing(MonthsToAdd) Then
      MonthsToAdd = 0
   End If

   TotalMonths = month(InputDate) + MonthsToAdd
   NewMonth = TotalMonths - (12 * Int(TotalMonths / 12))
   NewYear = Year(InputDate) + Int(TotalMonths / 12)

   If NewMonth = 0 Then
      NewMonth = 12
      NewYear = NewYear - 1
   End If

   Select Case NewMonth
      Case 1, 3, 5, 7, 8, 10, 12
         EOMONTH = DateSerial(NewYear, NewMonth, 31)
      Case 4, 6, 9, 11
         EOMONTH = DateSerial(NewYear, NewMonth, 30)
      Case 2
         If Int(NewYear / 4) = NewYear / 4 Then
            EOMONTH = DateSerial(NewYear, NewMonth, 29)
         Else
            EOMONTH = DateSerial(NewYear, NewMonth, 28)
         End If
   End Select
End Function
