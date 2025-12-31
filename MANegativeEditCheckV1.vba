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
Set outWB = Workbooks.Add(1) 'create a new one-worksheet workbook
exceloutfile = sPath & "\MA Negative Database Input " & WorksheetFunction.Text(Date, "mm-dd-yyyy") & ".xlsx"
outWB.SaveAs FileName:=exceloutfile, FileFormat:=51
Set dbws = ActiveSheet
writerow = 1

dbws.Range("A1") = "ReviewNo"
dbws.Range("B1") = "SampleMonth"
dbws.Range("C1") = "ReviewerNo"
dbws.Range("D1") = "StateCode"
dbws.Range("E1") = "CountyCode"
dbws.Range("F1") = "CaseNo"
dbws.Range("G1") = "CaseCategoryCode"
dbws.Range("H1") = "ProgramStatus"
dbws.Range("I1") = "GrantGroup"
dbws.Range("J1") = "AgencyDecisionDate"
dbws.Range("K1") = "AgencyActionDate"
dbws.Range("L1") = "ReviewCategory"
dbws.Range("M1") = "ActionTypeCode"
dbws.Range("N1") = "HearingReqCode"
dbws.Range("O1") = "ReasonForActionCode"
dbws.Range("P1") = "EligibilityRequirementCode"
dbws.Range("Q1") = "FieldInvestigationCode"
dbws.Range("R1") = "DispositionCode"
dbws.Range("S1") = "PostReviewStatusCode"

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
    
        Case "8"
            program = "MA Negative"
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
            '& "Review Month " & mname & " " & Left(monthstr, 4) & "\"

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
                    Workbooks.Open UpdateLinks:=0, FileName:=.FoundFiles(1), ReadOnly:=True
                    Set inWB = ActiveWorkbook
                    
                    writerow = writerow + 1
                    'Format row in write spreadsheet
                    dbws.Range("A" & writerow).NumberFormat = "@"
                    dbws.Range("B" & writerow).NumberFormat = "mm/dd/yyyy"
                    dbws.Range("C" & writerow & ":I" & writerow).NumberFormat = "@"
                    dbws.Range("J" & writerow & ":K" & writerow).NumberFormat = "m/d/yyyy"
                    dbws.Range("L" & writerow & ":S" & writerow).NumberFormat = "@"
                    
                    'Copy data into first sheet
                    dbws.Range("A" & writerow) = Sheets(reviewtxt).Range("L15") 'review number
                    dbws.Range("B" & writerow) = DateValue(Val(Left(Sheets(reviewtxt).Range("T11"), 2)) & "/1/" & Val(Right(Sheets(reviewtxt).Range("T11"), 4))) 'sample month
                    dbws.Range("C" & writerow) = Sheets(reviewtxt).Range("AB11") & Sheets(reviewtxt).Range("AC11") 'reviewer number
                    dbws.Range("D" & writerow) = Sheets(reviewtxt).Range("B15") 'state code
                    dbws.Range("E" & writerow) = Sheets(reviewtxt).Range("G15") 'county
                    'Check if case number has an A at the beginning
                    If Left(Sheets(reviewtxt).Range("S15"), 1) = "A" Then
                        dbws.Range("F" & writerow) = Mid(Sheets(reviewtxt).Range("S15"), 2) 'case number
                    Else
                        dbws.Range("F" & writerow) = Sheets(reviewtxt).Range("S15") 'case number
                    End If
                    dbws.Range("G" & writerow) = Trim(Sheets(reviewtxt).Range("AB15")) 'category
                    dbws.Range("H" & writerow) = Sheets(reviewtxt).Range("B19") 'program status
                    dbws.Range("I" & writerow) = Sheets(reviewtxt).Range("F19") 'grant group
                    
                    'Check if data fields have a date in them
                    If IsDate(Sheets(reviewtxt).Range("J19")) Then
                        dbws.Range("J" & writerow) = Sheets(reviewtxt).Range("J19") 'agency decision date
                    Else
                        dbws.Range("J" & writerow) = ""
                    End If
                    If IsDate(Sheets(reviewtxt).Range("S19")) Then
                        dbws.Range("K" & writerow) = Sheets(reviewtxt).Range("S19") 'agency action date
                    Else
                        dbws.Range("K" & writerow) = ""
                    End If
                    
                    dbws.Range("L" & writerow) = Sheets(reviewtxt).Range("AB19") 'review category
                    dbws.Range("M" & writerow) = Sheets(reviewtxt).Range("AF19") 'type of action
                    
                    'For these fields, if field is blank put B or BB in database field
                    If Sheets(reviewtxt).Range("C25") = "" Or Len(Trim(Sheets(reviewtxt).Range("C25"))) = 0 _
                        Or InStr(Sheets(reviewtxt).Range("C25"), "-") > 0 Then
                        dbws.Range("N" & writerow) = "B"
                    Else
                        dbws.Range("N" & writerow) = Sheets(reviewtxt).Range("C25")
                    End If
                    If Sheets(reviewtxt).Range("Q25") = "" Or Len(Trim(Sheets(reviewtxt).Range("Q25"))) = 0 _
                        Or InStr(Sheets(reviewtxt).Range("Q25"), "-") > 0 Then
                        dbws.Range("O" & writerow) = "BB"
                    Else
                        dbws.Range("O" & writerow) = Sheets(reviewtxt).Range("Q25")
                    End If
                    If Sheets(reviewtxt).Range("C40") = "" Or Len(Trim(Sheets(reviewtxt).Range("C40"))) = 0 _
                        Or InStr(Sheets(reviewtxt).Range("C40"), "-") > 0 Then
                        dbws.Range("P" & writerow) = "B"
                    Else
                        dbws.Range("P" & writerow) = Sheets(reviewtxt).Range("C40")
                    End If
                    If Sheets(reviewtxt).Range("C56") = "" Or Len(Trim(Sheets(reviewtxt).Range("C56"))) = 0 _
                        Or InStr(Sheets(reviewtxt).Range("C56"), "-") > 0 Then
                        dbws.Range("Q" & writerow) = "B"
                    Else
                        dbws.Range("Q" & writerow) = Sheets(reviewtxt).Range("C56")
                    End If
                    If Sheets(reviewtxt).Range("M56") = "" Or Len(Trim(Sheets(reviewtxt).Range("M56"))) = 0 _
                        Or InStr(Sheets(reviewtxt).Range("M56"), "-") > 0 Then
                        dbws.Range("R" & writerow) = "B"
                    Else
                        dbws.Range("R" & writerow) = Sheets(reviewtxt).Range("M56")
                    End If
                    If Sheets(reviewtxt).Range("Y56") = "" Or Len(Trim(Sheets(reviewtxt).Range("Y56"))) = 0 _
                        Or InStr(Sheets(reviewtxt).Range("Y56"), "-") > 0 Then
                        dbws.Range("S" & writerow) = "B"
                    Else
                        dbws.Range("S" & writerow) = Sheets(reviewtxt).Range("Y56")
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

    'Store excel worksheet into field office database
    'Copy blank database into local directory
    SrceFile = DLetter & "\HQ - Data Entry\Create FO Databases\FO Databases\MA_Neg_Blank.mdb"
    databasename = sPath & "\MA NEG1 " & WorksheetFunction.Text(Date, "mm-dd-yyyy") & ".mdb"
    
    filenum = 1
    'Check if file exists, change name until file doesn't exist
    Do Until Len(Dir(databasename)) = 0
        filenum = filenum + 1
        databasename = sPath & "\MA NEG" & filenum & " " & _
            WorksheetFunction.Text(Date, "mmddyyyy") & ".mdb"
    Loop
    
    DestFile = databasename
    FileCopy SrceFile, DestFile
    
    'Connection string for database
    stCon = "Provider=Microsoft.Ace.OLEDB.12.0;" & _
    "Data Source=" & databasename & ";"
     'SQL code for data Insert to Access
    stSQL = "INSERT INTO CaseReview_dtl SELECT * FROM [Sheet1$] IN '" _
    & outWB.FullName & "' 'Excel 8.0;'"
     'set connection variable
    Set cnt = New ADODB.Connection
     'open connection to Access db and run the SQL
    With cnt
        .Open stCon
        .CursorLocation = adUseClient
        .Execute (stSQL)
    End With
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