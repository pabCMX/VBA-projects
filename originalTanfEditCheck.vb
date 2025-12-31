Public rswr As Long, qcwr As Long, plwr As Long, hiwr As Long, efwr As Long, mewr As Long, revidval As Long
Public inWB As Workbook, outWB As Workbook, inWS As Worksheet, outWS As Worksheet

Sub Find_Write_Database_Files()

'Finds review schedules, other documents and scanned documents
'Copies files to folder for sending to eGov

Dim lCount As Long
Dim maxrow As Integer, maxrowex As Integer
Dim program As String, exname As String, exnumstr As String
Dim JustFileName As String, PathStr As String
Dim monthstr As String, reviewtxt As String
Dim wb As Workbook
Dim DUNC As String
Dim fs As clFileSearchModule
Dim outWB As Workbook
Dim cnt As ADODB.Connection
Dim stSQL As String, stCon As String, stDB As String

Set fs = New clFileSearchModule
                                      
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
    End If
Next i

If DLetter = "" Then
    MsgBox "Network Drive to Examiner Files are NOT correct" & Chr(13) & _
        "Contact Valerie or Wes"
    End
End If

pathdir = DLetter & "\Schedules by Examiner Number\"
'pathdir_scan = DLetter & "\Scanning for FNS\"

If Len(Dir(pathdir, vbDirectory)) = 0 Then
    MsgBox "Path to Examiner's File: " & pathdir & " does NOT exists!!" & Chr(13) & _
        "Contact Valerie or Wes"
    End
End If
' Keep Screen from updating
Application.ScreenUpdating = False

' Save Display Status Bar
'oldStatusBar = Application.DisplayStatusBar
Application.DisplayStatusBar = True

' Save name of the search for files workbook and worksheet
Set thissht = ActiveSheet
sPath = ActiveWorkbook.Path

' Find maximum row for cases
thissht.Range("E1").End(xlDown).Select
maxrow = ActiveCell.Row

' Find maximum row for examiners
thissht.Range("L1").End(xlDown).Select
maxrowex = ActiveCell.Row

'Open new excel workbook to write data into
'Copy blank excel workbook with proper headers into local directory
    SrceFile = sPath & "\FO Databases\TANF_Template.xlsx"
    exceloutfile = sPath & "\TANF Database Input " & WorksheetFunction.Text(Date, "mm-dd-yyyy") & ".xlsx"
    FileCopy SrceFile, exceloutfile

'Open excel template
Workbooks.Open FileName:=exceloutfile
Set outWB = ActiveWorkbook

'set write row varables for the tables
rswr = 1
qcwr = 1
plwr = 1
hiwr = 1
efwr = 1
eawr = 1


' Loop through all review numbers in workbook
For i = 2 To maxrow

    ' Set review number
    reviewtxt = WorksheetFunction.Text(thissht.Range("E" & i), "#")
    
    'Remove leading zero
    If Left(reviewtxt, 1) = "0" Then
        reviewtxt = Right(reviewtxt, Len(reviewtxt) - 1)
    End If

    ' Display Status of processing
    frac = 100 * (i - 2) / (maxrow - 1)
    pct = Round(frac, 0)
    strTemp = "Processing Review Number " & reviewtxt & _
            " - " & pct & "% - " & i - 2 & "/" & maxrow - 1 & " done. Please be patient..."
    Application.StatusBar = strTemp

    ' Find examiner name
    exname = ""
    For j = 2 To maxrowex
        If thissht.Range("G" & i) = thissht.Range("L" & j) Then
            exname = thissht.Range("K" & j)
            Exit For
        End If
    Next j
    
    If exname = "" Then
    ' if examiner name not in list, present message
        MsgBox "No Examiner Name found for Review " & _
            reviewtxt & " and Examiner Number " & _
            thissht.Range("G" & i) & ". Please check Review Number."
    Else
    
    ' Format examiner number
    exnumstr = WorksheetFunction.Text(thissht.Range("G" & i), "00")
    If Left(exnumstr, 1) = "0" Then
        exnumstr = Right(exnumstr, 1)
    End If
    
    ' Find program
    Select Case Left(reviewtxt, 1)
    
        Case "1"
            program = "TANF"
        Case Else
            MsgBox "Review Number " & reviewtxt & " is not a TANF Review Number"
            Exit For
    End Select
    
    
    'Sample month
    monthstr = thissht.Range("F" & i)
    
    Select Case Right(monthstr, 2)
        Case "01"
            mname = "January"
        Case "02"
            mname = "February"
        Case "03"
            mname = "March"
        Case "04"
            mname = "April"
        Case "05"
            mname = "May"
        Case "06"
            mname = "June"
         Case "07"
            mname = "July"
         Case "08"
            mname = "August"
         Case "09"
            mname = "September"
         Case "10"
            mname = "October"
         Case "11"
            mname = "November"
         Case "12"
            mname = "December"
    End Select
       
    ' Build Path to workbooks containing schedules and other worksheets
    PathStr = pathdir & exname & " - " & _
            exnumstr & "\" & program & "\"
            '& "Review Month " & _ mname & " " & Left(monthstr, 4) & "\"

    'starttime = Now()
            
    ' Search for file in the path
    With fs
        .NewSearch
        .SearchSubFolders = True
        
         ' Set path name
        .LookIn = PathStr
        .FileType = msoFileTypeExcelWorkbooks
        
        ' Set file name
        'Review Number 50002 Month 200910 Examiner 47.xls
        .FileName = "Review Number " & reviewtxt & " Month " & monthstr & _
            " Examiner" & "*.xls*"
         '   " Examiner " & exnumstr & ".xls"

            ' If file is found then start the copy process
            If .Execute > 0 Then 'Workbooks in folder
                'For lCount = 1 To .FoundFiles.Count 'Loop through all.

                    ' Open workbook with schedule
                    'Workbooks.Open FileName:=.FoundFiles(lCount), ReadOnly:=True
                    Workbooks.Open UpdateLinks:=0, FileName:=.FoundFiles(1), ReadOnly:=True
                    Set inWB = ActiveWorkbook
                    Set inWS = inWB.Sheets(reviewtxt)
                    disp_code = inWS.Range("AI10") 'Disposition Code
                    
                    revidval = i - 1
                    Set outWS = outWB.Sheets("Review_Summary_dtl")
                    Call revsum(outWS)
                    
                    'Only fill in these tables for completed cases
                    If disp_code = 1 Then
                        Set outWS = outWB.Sheets("QC_Case_Info_dtl")
                        Call qcinfo(outWS)
                        Set outWS = outWB.Sheets("Person_Level_Info_dtl")
                        Call plinfo(outWS)
                        Set outWS = outWB.Sheets("Household_Income_dtl")
                        Call hhinc(outWS)
                        Set outWS = outWB.Sheets("Error_Findings_dtl")
                        Call errfind(outWS)
                    End If
                    
                    'Close schedule workbook without saving
                    inWB.Close False
                    
                 'Next lCount
            Else
                    ' If schedule not found, present message
                    MsgBox "Review Schedule not found in examiner folder for review number " & thissht.Range("E" & i)
            End If ' File search result
    End With ' Search for file
  End If ' No examiner found
Next i ' Loop through review numbers

outWB.Save

    strTemp = "Storing results in Field Office Database"
    Application.StatusBar = strTemp

    'Store excel worksheet into field office database
    'Copy blank database into local directory
    SrceFile = sPath & "\FO Databases\TANF_Blank.mdb"
    databasename = sPath & "\TANF1 " & WorksheetFunction.Text(Date, "mm-dd-yyyy") & ".mdb"
    
    filenum = 1
    'Check if file exists, change name until file doesn't exist
    Do Until Len(Dir(databasename)) = 0
        filenum = filenum + 1
        databasename = sPath & "\TANF" & filenum & " " & _
            WorksheetFunction.Text(Date, "mmddyyyy") & ".mdb"
    Loop
    
    FileCopy SrceFile, databasename
    
    'Connection string for database
    stCon = "Provider=Microsoft.Ace.OLEDB.12.0;" & _
    "Data Source=" & databasename & ";"
    
    'Select worksheet name to be merged into database
    For n = 1 To 5
    Select Case n
        Case 1
            wsname = "Review_Summary_dtl"
            sheetrange = wsname & "$A1:O" & rswr
        Case 2
            wsname = "QC_Case_Info_dtl"
            sheetrange = wsname & "$A1:W" & qcwr
        Case 3
            wsname = "Person_Level_Info_dtl"
            sheetrange = wsname & "$A1:R" & plwr
        Case 4
            wsname = "Household_Income_dtl"
            sheetrange = wsname & "$A1:E" & hiwr
        Case 5
            wsname = "Error_Findings_dtl"
            sheetrange = wsname & "$A1:L" & efwr
    End Select
            
    'Check if there is data in the worksheet
    If outWB.Sheets(wsname).Range("A2") <> "" Then
         'SQL code for data Insert to Access
        stSQL = "INSERT INTO " & wsname & " SELECT * FROM [" & sheetrange & "] IN '" _
            & outWB.FullName & "' 'Excel 12.0 XML;'"
         'set connection variable
        Set cnt = New ADODB.Connection
         'open connection to Access db and run the SQL
        With cnt
            .Open stCon
            .CursorLocation = adUseClient
               .Execute (stSQL)
        End With
    End If
    Next n
     'close connection
    cnt.Close
     'release object from memory
    Set cnt = Nothing
    
    'Close excel workbook
    outWB.Close True
    'Delete excel workbook
    Kill exceloutfile
    
' Restore status bar
    Application.StatusBar = False

' Let Screen Update
    Application.ScreenUpdating = True
        
End Sub

Sub revsum(outWS As Worksheet)
    
                    rswr = rswr + 1
                    
                    'Copy data into first sheet
                    outWS.Range("A" & rswr) = revidval 'ReviewID
                    'Review Number
                    outWS.Range("B" & rswr) = inWS.Range("A10")
                    outWS.Range("C" & rswr) = inWS.Range("I10") 'Case Number
                    outWS.Range("D" & rswr) = inWS.Range("Q10") 'Category
                    outWS.Range("E" & rswr) = inWS.Range("S10") 'Grant Group
                    'Sample Month
                    outWS.Range("F" & rswr) = DateValue(Val(Left(inWS.Range("AB10"), 2)) & "/1/" & Val(Right(inWS.Range("AB10"), 4)))
                    
                    'if disposition is not 1, then make error amount blank
                    If inWS.Range("AI10") = 1 Then
                        outWS.Range("G" & rswr) = Round(Val(inWS.Range("AO10")) + 0.001, 0) 'Error Amount
                    Else
                        outWS.Range("G" & rswr) = "" 'for dropped cases, make blank
                    End If
                    
                    'Review Findings Code
                    If inWS.Range("AL10") = "" Or Len(Trim(inWS.Range("AL10"))) = 0 _
                        Or InStr(inWS.Range("AL10"), "-") > 0 Then
                        outWS.Range("H" & rswr) = "B"
                    Else
                        outWS.Range("H" & rswr) = inWS.Range("AL10")
                    End If
                    'District Code
                    If inWS.Range("Y10") = "" Or Len(Trim(inWS.Range("Y10"))) = 0 _
                        Or InStr(inWS.Range("Y10"), "-") > 0 Then
                        outWS.Range("I" & rswr) = "BB"
                    Else
                        outWS.Range("I" & rswr) = inWS.Range("Y10")
                    End If
                    
                    outWS.Range("J" & rswr) = inWS.Range("U10") 'CAO
                    outWS.Range("K" & rswr) = inWS.Range("AI10") 'Disposition Code
                    wsnaame = ""
                    
                    'Run Date
                    For Each ws In inWB.Worksheets
                        If ws.Name = "TANF Workbook" Then
                            wsname = ws.Name
                            Exit For
                        End If
                    Next ws
                    If wsname <> "" Then
                        outWS.Range("L" & rswr) = inWB.Sheets(wsname).Range("G33") 'Run Date
                    End If
                    
                    outWS.Range("M" & rswr) = inWS.Range("AO3") & inWS.Range("AP3") 'Examiner Number
                    'outWS.Range("N" & rswr) = inWS.Range("B16") 'Supplemental
                    
                    'Renewal Type
                    If inWS.Range("AB85") = "" Or Len(Trim(inWS.Range("AB85"))) = 0 _
                        Or InStr(inWS.Range("AB85"), "-") > 0 Then
                        outWS.Range("O" & rswr) = "B"
                    Else
                        outWS.Range("O" & rswr) = inWS.Range("AB85")
                    End If

End Sub

Sub qcinfo(outWS As Worksheet)
    
                    qcwr = qcwr + 1
                    
                    'Copy data into sheet
                    outWS.Range("A" & qcwr) = revidval 'ReviewID
                    outWS.Range("B" & qcwr) = inWS.Range("V20") 'Unborn child code
                    outWS.Range("C" & qcwr) = inWS.Range("Y20") 'Shelter Arrangement code
                    outWS.Range("D" & qcwr) = inWS.Range("J16") 'Prior Assistance code
                    outWS.Range("E" & qcwr) = inWS.Range("O20") 'Reason Protective Pay code
                    outWS.Range("F" & qcwr) = inWS.Range("U16") 'Action Type code
                    outWS.Range("G" & qcwr) = inWS.Range("A16") 'Most recent opening
                    outWS.Range("H" & qcwr) = inWS.Range("L16") 'Most recent action
                    outWS.Range("I" & qcwr) = Val(inWS.Range("W16")) 'Number of case members
                    outWS.Range("J" & qcwr) = inWS.Range("Z16") 'Liquid assests
                    outWS.Range("K" & qcwr) = inWS.Range("AE16") 'Real property
                    outWS.Range("L" & qcwr) = inWS.Range("AJ16") 'Countable Vehicle
                    outWS.Range("M" & qcwr) = inWS.Range("AO16") 'Non-Liquid Assests
                    outWS.Range("N" & qcwr) = inWS.Range("C20") 'Monthly Payments
                    outWS.Range("O" & qcwr) = inWS.Range("I20") 'Sample Month Payments
                    outWS.Range("P" & qcwr) = inWS.Range("Q20") 'Sanction Amount
                    outWS.Range("Q" & qcwr) = inWS.Range("AB20") 'Gross Income
                    outWS.Range("R" & qcwr) = inWS.Range("AH20") 'Income Disregard
                    outWS.Range("S" & qcwr) = inWS.Range("AN20") 'Net Income
                    outWS.Range("T" & qcwr) = inWS.Range("B24") 'FS Allotment
                    outWS.Range("U" & qcwr) = inWS.Range("U24") 'Over Payment Recoupment
                    outWS.Range("V" & qcwr) = inWS.Range("N20") 'Protective Payment Code
                    outWS.Range("W" & qcwr) = inWS.Range("AN24") 'TANF Days
                    
End Sub

Sub plinfo(outWS As Worksheet)
                            
            linenum = 0
            For j = 30 To 44 Step 2
                If inWS.Range("A" & j) = "" Then
                    Exit For
                Else
                    plwr = plwr + 1
                    
                    'Copy data into sheet
                    outWS.Range("A" & plwr) = revidval 'ReviewID
                    outWS.Range("B" & plwr) = inWS.Range("A" & j) 'Person number
                    outWS.Range("C" & plwr) = inWS.Range("D" & j) 'Case_Afl_First_Code
                    outWS.Range("D" & plwr) = inWS.Range("E" & j) 'Case_Afl_Second_Code
                    outWS.Range("E" & plwr) = inWS.Range("G" & j) 'Deprivation_Code
                    outWS.Range("F" & plwr) = inWS.Range("J" & j) 'Relationship_Payment_Code
                    outWS.Range("G" & plwr) = Val(inWS.Range("M" & j)) 'Age
                    outWS.Range("H" & plwr) = inWS.Range("P" & j) 'Gender
                    outWS.Range("I" & plwr) = inWS.Range("R" & j) 'Race
                    outWS.Range("J" & plwr) = inWS.Range("T" & j) 'Citizenship
                    outWS.Range("K" & plwr) = inWS.Range("V" & j) 'Education Level
                    outWS.Range("L" & plwr) = inWS.Range("Y" & j) 'Reset Code
                    outWS.Range("M" & plwr) = inWS.Range("AC" & j) 'Work Activity Code
                    outWS.Range("N" & plwr) = inWS.Range("AF" & j) 'Referral Days
                    outWS.Range("O" & plwr) = inWS.Range("AK" & j) 'Marital Status Code
                    outWS.Range("P" & plwr) = inWS.Range("AN" & j) 'Program Status Code
                    outWS.Range("Q" & plwr) = inWS.Range("AQ" & j) 'TE Code
                    linenum = linenum + 1
                    outWS.Range("R" & plwr) = linenum 'Line No
                
                End If
            Next j
                    
End Sub

Sub hhinc(outWS As Worksheet)
                            
            For j = 50 To 56 Step 2
                If inWS.Range("C" & j) = "" Then
                    Exit For
                Else
                    For k = 7 To 37 Step 10
                        If inWS.Cells(j, k) = "" Then
                            Exit For
                        Else
                            hiwr = hiwr + 1
                    
                            'Copy data into sheet
                            outWS.Range("A" & hiwr) = revidval 'ReviewID
                            outWS.Range("B" & hiwr) = inWS.Cells(j, k + 4) 'Amount of income
                            'not used outWS.Range("C" & plwr) = inWS.Range("F" & j)
                            outWS.Range("D" & hiwr) = inWS.Range("C" & j) 'Person number
                            outWS.Range("E" & hiwr) = inWS.Cells(j, k) 'Type of income
                        End If
                   Next k
                End If
            Next j
                    
End Sub


Sub errfind(outWS As Worksheet)
                            
            linenum = 0
            For j = 61 To 67 Step 2
                If inWS.Range("F" & j) = "" Then
                    Exit For
                Else
                    efwr = efwr + 1
                    
                    'Copy data into sheet
                    outWS.Range("A" & efwr) = revidval 'ReviewID
                    outWS.Range("B" & efwr) = inWS.Range("F" & j) 'Error_Findings Code
                    outWS.Range("C" & efwr) = inWS.Range("AR" & j) 'Occurrence_Period_Code
                    outWS.Range("D" & efwr) = inWS.Range("AD" & j) 'Discovery_Code
                    outWS.Range("E" & efwr) = inWS.Range("AH" & j) 'Verification_Code
                    outWS.Range("F" & efwr) = inWS.Range("T" & j) 'Client_Agency_Code
                    outWS.Range("G" & efwr) = inWS.Range("C" & j) 'Optional
                    outWS.Range("H" & efwr) = inWS.Range("X" & j) 'Dollar_Amount
                    'Occurence Date
                    If inWS.Range("AL" & j) <> "" Then
                        occmonth = Month(inWS.Range("AL" & j))
                        occyear = Year(inWS.Range("AL" & j))
                        outWS.Range("I" & efwr) = DateSerial(occyear, occmonth, 1)
                    End If
                    outWS.Range("J" & efwr) = inWS.Range("O" & j) 'Nature_Code
                    outWS.Range("K" & efwr) = inWS.Range("J" & j) 'Element_Code
                    linenum = linenum + 1
                    outWS.Range("L" & efwr) = linenum 'Line number
                
                End If
            Next j
                    
End Sub