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
    SrceFile = sPath & "\FO Databases\MA_Positive_Template.xlsx"
    exceloutfile = sPath & "\MA Positive Database Input " & WorksheetFunction.Text(Date, "mm-dd-yyyy") & ".xlsx"
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
mewr = 1


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
    
        Case "2"
            program = "MA Positive"
        Case Else
            MsgBox "Review Number " & reviewtxt & " is not a MA Negative Review Number"
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
            '& "Review Month " & _
            'mname & " " & Left(monthstr, 4) & "\"
            
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
                    disp_code = inWS.Range("F16") 'Disposition Code
                    
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
                        If inWS.Range("S16") <> 1 Then
                            Set outWS = outWB.Sheets("Error_Findings_dtl")
                            Call errfind(outWS)
                        End If
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
    SrceFile = sPath & "\FO Databases\MA_Pos_Blank.mdb"
    databasename = sPath & "\MA POS1 " & WorksheetFunction.Text(Date, "mm-dd-yyyy") & ".mdb"
    
    filenum = 1
    'Check if file exists, change name until file doesn't exist
    Do Until Len(Dir(databasename)) = 0
        filenum = filenum + 1
        databasename = sPath & "\MA POS" & filenum & " " & _
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
            sheetrange = wsname & "$A1:Z" & rswr
        Case 2
            wsname = "QC_Case_Info_dtl"
            sheetrange = wsname & "$A1:L" & qcwr
        Case 3
            wsname = "Person_Level_Info_dtl"
            sheetrange = wsname & "$A1:N" & plwr
        Case 4
            wsname = "Household_Income_dtl"
            sheetrange = wsname & "$A1:E" & hiwr
        Case 5
            wsname = "Error_Findings_dtl"
            sheetrange = wsname & "$A1:N" & efwr
    End Select
            
    'Check if there is data in the worksheet
    If outWB.Sheets(wsname).Range("A2") <> "" Then
         'SQL code for data Insert to Access
        stSQL = "INSERT INTO " & wsname & " SELECT * FROM [" & sheetrange & "] IN '" _
            & outWB.FullName & "' 'Excel 8.0;'"
         'outWB.Sheets(wsname).Range("A10") = stSQL
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
'fill in review summary from section I of schedule
                    rswr = rswr + 1
                    
                    'Copy data into first sheet
                    outWS.Range("A" & rswr) = revidval 'ReviewID
                    'Review Number
                    outWS.Range("B" & rswr) = inWS.Range("A10") & inWS.Range("B10") & inWS.Range("C10") & inWS.Range("D10") & inWS.Range("E10") & inWS.Range("F10")
                    'Managed Care
                    If inWS.Range("AK10") = "" Or Len(Trim(inWS.Range("AK10"))) = 0 _
                        Or InStr(inWS.Range("AK10"), "-") > 0 Then
                        outWS.Range("C" & rswr) = "B"
                    Else
                        outWS.Range("C" & rswr) = inWS.Range("AK10")
                    End If
                    outWS.Range("D" & rswr) = inWS.Range("H10") 'Case Number
                    outWS.Range("E" & rswr) = inWS.Range("Q10") 'Category
                    outWS.Range("F" & rswr) = inWS.Range("V10") 'Program Status
                    outWS.Range("G" & rswr) = inWS.Range("Z10") 'Grant Group
                    'Sample Month
                    outWS.Range("H" & rswr) = DateValue(Val(Left(inWS.Range("AM10"), 2)) & "/1/" & Val(Right(inWS.Range("AM10"), 4)))
                    'District Code
                    If inWS.Range("AH10") = "" Or Len(Trim(inWS.Range("AH10"))) = 0 _
                        Or InStr(inWS.Range("AH10"), "-") > 0 Then
                        outWS.Range("I" & rswr) = "BB"
                    Else
                        outWS.Range("I" & rswr) = inWS.Range("AH10")
                    End If
                    outWS.Range("J" & rswr) = inWS.Range("AC10") 'CAO
                    outWS.Range("K" & rswr) = inWS.Range("F16") 'Disposition Code
                    wsnaame = ""
                    'Run Date
                    For Each ws In inWB.Worksheets
                        If ws.Name = "MA Workbook" Or ws.Name = "MA Facesheet" Then
                            wsname = ws.Name
                            Exit For
                        End If
                    Next ws
                    If wsname <> "" Then
                        outWS.Range("L" & rswr) = inWB.Sheets(wsname).Range("H35") 'Run Date
                    End If
                    outWS.Range("M" & rswr) = inWS.Range("AO3") & inWS.Range("AP3") 'Examiner Number
                    outWS.Range("N" & rswr) = inWS.Range("B16") 'Stratum
                    outWS.Range("O" & rswr) = inWS.Range("I16") 'Elig Cov Agency
                    outWS.Range("P" & rswr) = Trim(inWS.Range("S16")) 'Initial Eligibility
                    outWS.Range("Q" & rswr) = Trim(inWS.Range("AF16")) 'Excess Resource Amount
                    outWS.Range("R" & rswr) = Trim(inWS.Range("AN16")) 'Paid claims
                    'Final eligibility
                    If inWS.Range("B21") = "" Or Len(Trim(inWS.Range("B21"))) = 0 _
                        Or InStr(inWS.Range("B21"), "-") > 0 Then
                        outWS.Range("S" & rswr) = "BB"
                    Else
                        outWS.Range("S" & rswr) = Trim(inWS.Range("B21"))
                    End If
                    outWS.Range("T" & rswr) = Trim(inWS.Range("H21")) 'Revised Initial
                    outWS.Range("U" & rswr) = Trim(inWS.Range("Q21")) 'Spend Down
                    outWS.Range("V" & rswr) = Trim(inWS.Range("V21")) 'Initial LU Errors
                    outWS.Range("W" & rswr) = Trim(inWS.Range("AE21")) 'Liability Errors
                    outWS.Range("X" & rswr) = Trim(inWS.Range("AN21")) 'Eligibility Errors
                    outWS.Range("Y" & rswr) = Trim(inWS.Range("N16")) 'Eligibility Coverage QC
                    outWS.Range("Z" & rswr) = Trim(inWS.Range("X16")) 'Initial Liability Error
End Sub

Sub qcinfo(outWS As Worksheet)
'fill in qc info from section II
                    qcwr = qcwr + 1
                    
                    'Copy data into sheet
                    outWS.Range("A" & qcwr) = revidval 'ReviewID
                    outWS.Range("B" & qcwr) = inWS.Range("J27") 'Prior assistance code
                    outWS.Range("C" & qcwr) = inWS.Range("V27") 'Action Type code
                    outWS.Range("D" & qcwr) = inWS.Range("A27") 'Most recent opening
                    outWS.Range("E" & qcwr) = inWS.Range("L27") 'Most recent action
                    outWS.Range("F" & qcwr) = Val(Trim(inWS.Range("Y27"))) 'Number of case members
                    outWS.Range("G" & qcwr) = Trim(inWS.Range("AB27")) 'Liquid assests
                    outWS.Range("H" & qcwr) = Trim(inWS.Range("AH27")) 'Real property
                    outWS.Range("I" & qcwr) = Trim(inWS.Range("AN27")) 'Countable Vehicle
                    outWS.Range("J" & qcwr) = Trim(inWS.Range("B32")) 'Non-Liquid Assests
                    outWS.Range("K" & qcwr) = Trim(inWS.Range("AF32")) 'Gross Income
                    outWS.Range("L" & qcwr) = Trim(inWS.Range("AN32")) 'Net Income
                    
End Sub

Sub plinfo(outWS As Worksheet)
'fill in person data from section III
            linenum = 0
            For j = 51 To 73 Step 2
                If inWS.Range("B" & j) = "" Then
                    Exit For
                Else
                    plwr = plwr + 1
                    
                    'Copy data into sheet
                    outWS.Range("A" & plwr) = revidval 'ReviewID
                    outWS.Range("B" & plwr) = Val(inWS.Range("B" & j)) 'Person number
                    outWS.Range("C" & plwr) = inWS.Range("F" & j) & inWS.Range("G" & j) 'FS Case Aff
                    outWS.Range("D" & plwr) = inWS.Range("J" & j) 'TANF/MA Case Aff
                    outWS.Range("E" & plwr) = inWS.Range("N" & j) 'Relationship to HH
                    outWS.Range("F" & plwr) = Trim(inWS.Range("R" & j)) 'Age
                    outWS.Range("G" & plwr) = inWS.Range("V" & j) 'Gender
                    outWS.Range("H" & plwr) = inWS.Range("Y" & j) 'Race
                    outWS.Range("I" & plwr) = inWS.Range("AB" & j) 'Citizenship
                    outWS.Range("J" & plwr) = inWS.Range("AF" & j) 'Education Level
                    'Employment Training status
                    If inWS.Range("AI" & j) = "" Or Len(Trim(inWS.Range("AI" & j))) = 0 _
                        Or InStr(inWS.Range("AI" & j), "-") > 0 Then
                        outWS.Range("K" & plwr) = "BB"
                    Else
                        outWS.Range("K" & plwr) = inWS.Range("AI" & j)
                    End If
                    'Employment status
                    If inWS.Range("AM" & j) = "" Or Len(Trim(inWS.Range("AM" & j))) = 0 _
                        Or InStr(inWS.Range("AM" & j), "-") > 0 Then
                        outWS.Range("L" & plwr) = "BB"
                    Else
                        outWS.Range("L" & plwr) = inWS.Range("AM" & j)
                    End If
                    outWS.Range("M" & plwr) = inWS.Range("AQ" & j) 'Institutional Status
                    linenum = linenum + 1
                    outWS.Range("N" & plwr) = linenum 'Line number
                
                End If
            Next j
                    
End Sub

Sub hhinc(outWS As Worksheet)
'fill in income from section IV

            For j = 78 To 84 Step 2
                If inWS.Range("B" & j) = "" Then
                    Exit For
                Else
                    For k = 6 To 36 Step 10
                        If inWS.Cells(j, k) = "" Then
                            Exit For
                        Else
                            hiwr = hiwr + 1
                    
                            'Copy data into sheet
                            outWS.Range("A" & hiwr) = revidval 'ReviewID
                            outWS.Range("B" & hiwr) = Trim(inWS.Cells(j, k + 4)) 'Amount of income
                            'not used outWS.Range("C" & plwr) = inWS.Range("F" & j)
                            outWS.Range("D" & hiwr) = Val(inWS.Range("B" & j)) 'Person number
                            outWS.Range("E" & hiwr) = inWS.Cells(j, k) 'Type of income
                        End If
                   Next k
                End If
            Next j
                    
End Sub


Sub errfind(outWS As Worksheet)
'fill in error findings from section V
            linenum = 0
            For j = 96 To 112 Step 2
                If inWS.Range("B" & j) = "" Then
                    Exit For
                Else
                    efwr = efwr + 1
                    
                    'Copy data into sheet
                    outWS.Range("A" & efwr) = revidval 'ReviewID
                    outWS.Range("B" & efwr) = inWS.Range("B" & j) 'Program ID
                    outWS.Range("C" & efwr) = inWS.Range("D" & j) 'Error_Findings_Eligibility
                    'Error_Findings_Code
                    If inWS.Range("G" & j) = "" Or Len(Trim(inWS.Range("G" & j))) = 0 _
                        Or InStr(inWS.Range("G" & j), "-") > 0 Then
                        outWS.Range("D" & efwr) = "BB"
                    Else
                        outWS.Range("D" & efwr) = inWS.Range("G" & j)
                    End If
                    outWS.Range("E" & efwr) = inWS.Range("K" & j) 'Error Member
                    outWS.Range("F" & efwr) = inWS.Range("O" & j) 'Element Code
                    outWS.Range("G" & efwr) = inWS.Range("T" & j) 'Nature Code
                    outWS.Range("H" & efwr) = inWS.Range("X" & j) 'Client Agency Code
                    outWS.Range("I" & efwr) = Trim(inWS.Range("AA" & j)) 'Dollar Amount
                    outWS.Range("J" & efwr) = inWS.Range("AG" & j) 'Discovery Code
                    outWS.Range("K" & efwr) = inWS.Range("AJ" & j) 'Verification Code
                    'Occurence Date
                    If inWS.Range("AL" & j) <> "" Then
                        occmonth = Month(inWS.Range("AL" & j))
                        occyear = Year(inWS.Range("AL" & j))
                        outWS.Range("L" & efwr) = DateSerial(occyear, occmonth, 1)
                    End If
                    outWS.Range("M" & efwr) = inWS.Range("AQ" & j) 'Occurence Period Code
                    linenum = linenum + 1
                    outWS.Range("N" & efwr) = linenum 'Line number
                
                End If
            Next j
                    
End Sub