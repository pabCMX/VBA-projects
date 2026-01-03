Attribute VB_Name = "CAO_Appointment"
Public ApptDate As String, ApptTime As String

Sub CAOAppt()

'Create CAO Appointment Memo
Dim appWD As Object
Dim review_number As Long, TempStr As String
Dim sample_month As String, PathStr As String, findmemo As String
Dim countynum As String, districtnum As String

'Find path to Examiner's files on this PC
Set WshNetwork = CreateObject("WScript.Network")
Set oDrives = WshNetwork.EnumNetworkDrives

Dim DLetter As String, DUNC As String
Dim thisws As Worksheet, thiswb As Workbook
Dim sPath As String, tempsourcename As String

sPath = ActiveWorkbook.Path

Set thisws = ActiveSheet
Set thiswb = ActiveWorkbook

'picking date of appointment
SelectDate.Show
'picking time of appointment
SelectTime.Show

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

PathStr = DLetter
Application.ScreenUpdating = False

Application.StatusBar = "Starting mail merge ..."

'Copy template data source to active directory
tempsourcename = sPath & "\FM DS Temp.xlsx"
SrceFile = PathStr & "Finding Memo\Finding Memo Data Source.xlsx"
DestFile = tempsourcename
FileCopy SrceFile, DestFile

'Open spreadsheet where all the data is stored to populate findings memo
Workbooks.Open fileName:=tempsourcename, UpdateLinks:=False

Dim datasourcewb As Workbook, datasourcews As Worksheet

Set datasourcewb = ActiveWorkbook
Set datasourcews = ActiveSheet

'deleting information in Finding Memo Data Source spreadsheet
Range("B2:E2") = ""
Range("G2") = ""
Range("O2:S2") = ""
Range("W2:Y2") = ""

'Enter appointment date and time into data source file
Range("AR2") = ApptDate
Range("AS2") = ApptTime

'determing what kind of review it is by the first number in the review
review_type = Left(thisws.Name, 1)

'SNAP Positive
If review_type = 5 Then

'Client Information
Range("B2") = StrConv(thisws.Range("B4"), vbProperCase)
Range("C2") = thisws.Range("C155") & thisws.Range("D155") & "/" & Left(thisws.Range("I18"), 9)

' Get county number and name in correct format
countynum = thisws.Range("C155") & thisws.Range("D155")
districtnum = thisws.Range("E155") & thisws.Range("F155")
countylookup = countyformat(countynum, thisws.Range("M5"), districtnum)
Range("AK2") = thisws.Range("M5") & " Assistance Office"

'Look for counties with districts
'If countynum = "02" Or countynum = "23" Or countynum = "40" Or _
'        countynum = "51" Or countynum = "63" Or countynum = "65" Then
'    Range("AK2") = thisws.Range("M5") & " AssisOffice"
'Else
'    Range("AK2") = thisws.Range("M5") & " Assistance Office"
'End If

'Put into temp data file other data
Range("D2") = thisws.Range("A18")
review_number = Range("D2")
Range("E2") = thisws.Range("AD18") & "/" & thisws.Range("AG18")
sample_month = thisws.Range("AD18") & thisws.Range("AG18")
Range("A7") = Val(thisws.Range("AJ5") & thisws.Range("AK5"))
Range("G2") = Range("C7") & "/" & thisws.Range("AE5")
Range("E7") = review_type
Range("O2") = thiswb.Sheets("FS Computation").Range("B60")
Range("P2") = thiswb.Sheets("FS Computation").Range("C60")
Range("Q2") = thisws.Range("Y22")
Range("R2") = thisws.Range("B29") & " - " & thisws.Range("G29") & " - " & thisws.Range("K29")
Range("E17") = Val(thisws.Range("K22"))

'reformat street address
If thisws.Range("B6") <> "" Then
    Range("Z2") = StrConv(thisws.Range("B5"), vbProperCase) & ", " & StrConv(thisws.Range("B6"), vbProperCase)
Else
    Range("Z2") = StrConv(thisws.Range("B5"), vbProperCase)
End If

'Reformat city and state
    temparry = Split(thisws.Range("B7"), ",")
    TempStr = StrConv(Trim(temparry(LBound(temparry))), vbProperCase)

Range("AA2") = TempStr & ", " & Trim(temparry(UBound(temparry)))

Else

    MsgBox ("CAO Appointment Letter not needed for this type of review")

End

End If

'Going out to website with list of Executive Directors
        FullTextFileName = datasourcews.Range("U6")
            
            Workbooks.Open fileName:=FullTextFileName, UpdateLinks:=False, ReadOnly:=True
            
            Range("A1:D1000").UnMerge

            'Allegheny Districts are formatted as dates, change them to 02/XX format
            For irow = 1 To 100
                If IsDate(Range("A" & irow)) Then
                    TempStr = WorksheetFunction.Text(Range("A" & irow), "mm/d/yyyy")
                    Range("A" & irow).NumberFormat = "@"
                    Range("A" & irow) = Left(TempStr, 4)
                End If
            Next irow
            
            'Look for county or district number
                found_row = Range("A1:A1000").Find(countylookup).Row
            
            datasourcews.Range("Y2") = Range("D" & found_row)
            'For irow = 1 To 1000
            '    If Range("A" & irow) = countylookup Then
                    'Get ED name
            '        datasourcews.Range("Y2") = Range("D" & irow)
            '        'Get CAO/DO Address
                    caoaddress = ""
                    For ii = 0 To 4
                        If Range("C" & found_row + ii) = "" Then
                            Exit For
                        End If
                        caoaddress = caoaddress & Range("C" & found_row + ii) & ", "
                    Next ii
                    datasourcews.Range("AL2") = caoaddress
                'End If
            'Next irow
            ActiveWorkbook.Close False
            
    'Put appointment time into Calendar
    Dim olApp As Object
    
    Application.DisplayAlerts = False

    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")

    If Err.Number = 429 Then
        Set olApp = CreateObject("Outlook.application")
    End If

    On Error GoTo 0

    Set olApt = olApp.CreateItem(1)

    With olApt
        temptime = TimeValue(WorksheetFunction.Text(datasourcews.Range("AS2"), "hh:mm AM/PM"))
        tempdate = WorksheetFunction.Text(datasourcews.Range("AR4"), "mm/dd/yyyy")
        .Start = DateValue(tempdate) + TimeValue(temptime)
        .End = .Start + TimeValue("01:00:00")
        .Subject = "Appointment with " & datasourcews.Range("B2")
        .Location = datasourcews.Range("AK2")
        .Body = "Review Number " & review_number & Chr(10) & _
            "Case Number " & Left(thisws.Range("I18"), 9)
        .BusyStatus = 3
        .ReminderMinutesBeforeStart = 5760
        .ReminderSet = True
        .SAVE
    End With

    Set olApt = Nothing
    Set olApp = Nothing
    
    datasourcewb.Close True
    Application.DisplayAlerts = True
            
    findmemo = "CAO Appointment Letter for Review Number " & review_number & " for Sample Month " & sample_month & ".doc"

    'Copy template cao appointment memo to active directory
    SrceFile = PathStr & "Finding Memo\CAO Appointment Letter Master.doc"
    tempfindmemoname = sPath & "\FM temp.doc"
    DestFile = tempfindmemoname
    FileCopy SrceFile, DestFile

    Set appWD = CreateObject("Word.Application")
    appWD.Visible = True

    With appWD
    .Documents.Open fileName:=tempfindmemoname
    
    'Merge CAO information into template
    With .ActiveDocument.MailMerge
        .OpenDataSource Name:=(tempsourcename), _
            ReadOnly:=True, LinkToSource:=0, AddToRecentFiles:=False, _
            PasswordDocument:="", PasswordTemplate:="", WritePasswordDocument:="", _
            WritePasswordTemplate:="", Revert:=False, Format:=0, _
            Connection:="yourNamedRange", SQLStatement:="SELECT * FROM `yourNamedRange`", SQLStatement1:=""
        .Destination = 0
        .SuppressBlankLines = True
        .Execute Pause:=False
    End With
    
Application.StatusBar = "Now creating document"

    .ActiveDocument.Content.Font.Name = "Times New Roman"
    '.ActiveDocument.Content.Font.Size = 12
    
    .ActiveDocument.SaveAs sPath & "\" & findmemo
    .ActiveDocument.Close
    
    .Documents(tempfindmemoname).Close SaveChanges:=False

End With
appWD.Quit
Set appWD = Nothing

'delete temp files
Kill tempfindmemoname
Kill tempsourcename

Application.StatusBar = False
Application.ScreenUpdating = True
Application.Visible = True

MsgBox "A CAO Appointment Letter has been saved as " & sPath & "\" & findmemo & _
". Please close the Internet Explorer with the CAO Directory. - Thank you!"

End Sub
Sub SpCAOAppt()

'Create CAO Appointment Memo
Dim appWD As Object
Dim review_number As Long, TempStr As String
Dim sample_month As String, PathStr As String, findmemo As String
Dim countynum As String, districtnum As String

'Find path to Examiner's files on this PC
Set WshNetwork = CreateObject("WScript.Network")
Set oDrives = WshNetwork.EnumNetworkDrives

Dim DLetter As String, DUNC As String
Dim thisws As Worksheet, thiswb As Workbook
Dim sPath As String, tempsourcename As String

sPath = ActiveWorkbook.Path

Set thisws = ActiveSheet
Set thiswb = ActiveWorkbook

'picking date of appointment
SelectDate.Show
'picking time of appointment
SelectTime.Show

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

PathStr = DLetter
Application.ScreenUpdating = False

Application.StatusBar = "Starting mail merge ..."

'Copy template data source to active directory
tempsourcename = sPath & "\FM DS Temp.xlsx"
SrceFile = PathStr & "Finding Memo\Finding Memo Data Source.xlsx"
DestFile = tempsourcename
FileCopy SrceFile, DestFile

'Open spreadsheet where all the data is stored to populate findings memo
Workbooks.Open fileName:=tempsourcename, UpdateLinks:=False

Dim datasourcewb As Workbook, datasourcews As Worksheet

Set datasourcewb = ActiveWorkbook
Set datasourcews = ActiveSheet

'deleting information in Finding Memo Data Source spreadsheet
Range("B2:E2") = ""
Range("G2") = ""
Range("O2:S2") = ""
Range("W2:Y2") = ""

'Enter appointment date and time into data source file
Range("AR2") = ApptDate
Range("AS2") = ApptTime

'determing what kind of review it is by the first number in the review
review_type = Left(thisws.Name, 1)

'SNAP Positive
If review_type = 5 Then

'Client Information
Range("B2") = StrConv(thisws.Range("B4"), vbProperCase)
Range("C2") = thisws.Range("C155") & thisws.Range("D155") & "/" & Left(thisws.Range("I18"), 9)

' Get county number and name in correct format
countynum = thisws.Range("C155") & thisws.Range("D155")
districtnum = thisws.Range("E155") & thisws.Range("F155")
countylookup = countyformat(countynum, thisws.Range("M5"), districtnum)
Range("AK2") = thisws.Range("M5") & " Assistance Office"

'Look for counties with districts
'If countynum = "02" Or countynum = "23" Or countynum = "40" Or _
'        countynum = "51" Or countynum = "63" Or countynum = "65" Then
'    Range("AK2") = thisws.Range("M5") & " AssisOffice"
'Else
'    Range("AK2") = thisws.Range("M5") & " Assistance Office"
'End If

'Put into temp data file other data
Range("D2") = thisws.Range("A18")
review_number = Range("D2")
Range("E2") = thisws.Range("AD18") & "/" & thisws.Range("AG18")
sample_month = thisws.Range("AD18") & thisws.Range("AG18")
Range("A7") = Val(thisws.Range("AJ5") & thisws.Range("AK5"))
Range("G2") = Range("C7") & "/" & thisws.Range("AE5")
Range("E7") = review_type
Range("O2") = thiswb.Sheets("FS Computation").Range("B60")
Range("P2") = thiswb.Sheets("FS Computation").Range("C60")
Range("Q2") = thisws.Range("Y22")
Range("R2") = thisws.Range("B29") & " - " & thisws.Range("G29") & " - " & thisws.Range("K29")
Range("E17") = Val(thisws.Range("K22"))

'reformat street address
If thisws.Range("B6") <> "" Then
    Range("Z2") = StrConv(thisws.Range("B5"), vbProperCase) & ", " & StrConv(thisws.Range("B6"), vbProperCase)
Else
    Range("Z2") = StrConv(thisws.Range("B5"), vbProperCase)
End If

'Reformat city and state
    temparry = Split(thisws.Range("B7"), ",")
    TempStr = StrConv(Trim(temparry(LBound(temparry))), vbProperCase)

Range("AA2") = TempStr & ", " & Trim(temparry(UBound(temparry)))

Else

    MsgBox ("CAO Appointment Letter not needed for this type of review")

End

End If

'Going out to website with list of Executive Directors
        FullTextFileName = datasourcews.Range("U6")
            
            Workbooks.Open fileName:=FullTextFileName, UpdateLinks:=False, ReadOnly:=True
            
            Range("A1:D1000").UnMerge

            'Allegheny Districts are formatted as dates, change them to 02/XX format
            For irow = 1 To 100
                If IsDate(Range("A" & irow)) Then
                    TempStr = WorksheetFunction.Text(Range("A" & irow), "mm/d/yyyy")
                    Range("A" & irow).NumberFormat = "@"
                    Range("A" & irow) = Left(TempStr, 4)
                End If
            Next irow
            
            'Look for county or district number
                found_row = Range("A1:A1000").Find(countylookup).Row
            
            datasourcews.Range("Y2") = Range("D" & found_row)
            'For irow = 1 To 1000
            '    If Range("A" & irow) = countylookup Then
                    'Get ED name
            '        datasourcews.Range("Y2") = Range("D" & irow)
            '        'Get CAO/DO Address
                    caoaddress = ""
                    For ii = 0 To 4
                        If Range("C" & found_row + ii) = "" Then
                            Exit For
                        End If
                        caoaddress = caoaddress & Range("C" & found_row + ii) & ", "
                    Next ii
                    datasourcews.Range("AL2") = caoaddress
                'End If
            'Next irow
            ActiveWorkbook.Close False
            
    'Put appointment time into Calendar
    Dim olApp As Object
    
    Application.DisplayAlerts = False

    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")

    If Err.Number = 429 Then
        Set olApp = CreateObject("Outlook.application")
    End If

    On Error GoTo 0

    Set olApt = olApp.CreateItem(1)

    With olApt
        temptime = TimeValue(WorksheetFunction.Text(datasourcews.Range("AS2"), "hh:mm AM/PM"))
        tempdate = WorksheetFunction.Text(datasourcews.Range("AR4"), "mm/dd/yyyy")
        .Start = DateValue(tempdate) + TimeValue(temptime)
        .End = .Start + TimeValue("01:00:00")
        .Subject = "Appointment with " & datasourcews.Range("B2")
        .Location = datasourcews.Range("AK2")
        .Body = "Review Number " & review_number & Chr(10) & _
            "Case Number " & Left(thisws.Range("I18"), 9)
        .BusyStatus = 3
        .ReminderMinutesBeforeStart = 5760
        .ReminderSet = True
        .SAVE
    End With

    Set olApt = Nothing
    Set olApp = Nothing
    
    datasourcewb.Close True
    Application.DisplayAlerts = True
            
    findmemo = "CAO Appointment Letter for Review Number " & review_number & " for Sample Month " & sample_month & ".doc"

    'Copy template cao appointment memo to active directory
    SrceFile = PathStr & "Finding Memo\Spanish CAO Appointment Letter.doc"
    tempfindmemoname = sPath & "\FM temp.doc"
    DestFile = tempfindmemoname
    FileCopy SrceFile, DestFile

    Set appWD = CreateObject("Word.Application")
    appWD.Visible = True

    With appWD
    .Documents.Open fileName:=tempfindmemoname
    
    'Merge CAO information into template
    With .ActiveDocument.MailMerge
        .OpenDataSource Name:=(tempsourcename), _
            ReadOnly:=True, LinkToSource:=0, AddToRecentFiles:=False, _
            PasswordDocument:="", PasswordTemplate:="", WritePasswordDocument:="", _
            WritePasswordTemplate:="", Revert:=False, Format:=0, _
            Connection:="yourNamedRange", SQLStatement:="SELECT * FROM `yourNamedRange`", SQLStatement1:=""
        .Destination = 0
        .SuppressBlankLines = True
        .Execute Pause:=False
    End With
    
Application.StatusBar = "Now creating document"

    .ActiveDocument.Content.Font.Name = "Times New Roman"
    '.ActiveDocument.Content.Font.Size = 12
    
    .ActiveDocument.SaveAs sPath & "\" & findmemo
    .ActiveDocument.Close
    
    .Documents(tempfindmemoname).Close SaveChanges:=False

End With
appWD.Quit
Set appWD = Nothing

'delete temp files
Kill tempfindmemoname
Kill tempsourcename

Application.StatusBar = False
Application.ScreenUpdating = True
Application.Visible = True

MsgBox "A Spanish CAO Appointment Letter has been saved as " & sPath & "\" & findmemo & _
". Please close the Internet Explorer with the CAO Directory. - Thank you!"

End Sub
Sub SpTeleAppt()

'Create Telephone Appointment memo
Dim appWD As Object
Dim review_number As Long
Dim sample_month As String, PathStr As String, findmemo As String
Dim countynum As String, districtnum As String

'Find path to Examiner's files on this PC
Set WshNetwork = CreateObject("WScript.Network")
Set oDrives = WshNetwork.EnumNetworkDrives

Dim DLetter As String, DUNC As String
Dim thisws As Worksheet, thiswb As Workbook
Dim sPath As String, tempsourcename As String

sPath = ActiveWorkbook.Path

Set thisws = ActiveSheet
Set thiswb = ActiveWorkbook

'picking date of appointment
SelectDate.Show
'picking time of appointment
SelectTime.Show

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

PathStr = DLetter
Application.ScreenUpdating = False

Application.StatusBar = "Starting mail merge ..."

'Copy template data source to active directory
tempsourcename = sPath & "\FM DS Temp.xlsx"
SrceFile = PathStr & "Finding Memo\Finding Memo Data Source.xlsx"
DestFile = tempsourcename
FileCopy SrceFile, DestFile

'Open spreadsheet where all the data is stored to populate findings memo
Workbooks.Open fileName:=tempsourcename, UpdateLinks:=False

Dim datasourcewb As Workbook, datasourcews As Worksheet

Set datasourcewb = ActiveWorkbook
Set datasourcews = ActiveSheet

'deleting information in Finding Memo Data Source spreadsheet
Range("B2:E2") = ""
Range("G2") = ""
Range("O2:S2") = ""
Range("W2:Y2") = ""

'Enter appointment date and time into data source file
Range("AR2") = ApptDate
Range("AS2") = ApptTime

'determing what kind of review it is by the first number in the review
review_type = Left(thisws.Name, 1)

'SNAP Positive
If review_type = 5 Then

'Client Information
Range("B2") = StrConv(thisws.Range("B4"), vbProperCase)
Range("C2") = thisws.Range("C155") & thisws.Range("D155") & "/" & Left(thisws.Range("I18"), 9)

' Get county number and name in correct format
countynum = thisws.Range("C155") & thisws.Range("D155")
districtnum = thisws.Range("E155") & thisws.Range("F155")
countylookup = countyformat(countynum, thisws.Range("M5"), districtnum)
Range("AK2") = thisws.Range("M5") & " Assistance Office"
'Look for counties with districts
'If countynum = "02" Or countynum = "23" Or countynum = "40" Or _
'        countynum = "51" Or countynum = "63" Or countynum = "65" Then
'    Range("AK2") = thisws.Range("M5") & " Office"
'Else
'    Range("AK2") = thisws.Range("M5") & " CAO"
'End If

'Put into temp data file other data
Range("D2") = thisws.Range("A18")
review_number = Range("D2")
Range("E2") = thisws.Range("AD18") & "/" & thisws.Range("AG18")
sample_month = thisws.Range("AD18") & thisws.Range("AG18")
Range("A7") = Val(thisws.Range("AJ5") & thisws.Range("AK5"))
Range("G2") = Range("C7") & "/" & thisws.Range("AE5")
Range("E7") = review_type
Range("O2") = thiswb.Sheets("FS Computation").Range("B60")
Range("P2") = thiswb.Sheets("FS Computation").Range("C60")
Range("Q2") = thisws.Range("Y22")
Range("R2") = thisws.Range("B29") & " - " & thisws.Range("G29") & " - " & thisws.Range("K29")
Range("E17") = Val(thisws.Range("K22"))
'reformat street address
If thisws.Range("B6") <> "" Then
    Range("Z2") = StrConv(thisws.Range("B5"), vbProperCase) & ", " & StrConv(thisws.Range("B6"), vbProperCase)
Else
    Range("Z2") = StrConv(thisws.Range("B5"), vbProperCase)
End If

'Reformat city and state
    temparry = Split(thisws.Range("B7"), ",")
    TempStr = StrConv(Trim(temparry(LBound(temparry))), vbProperCase)

Range("AA2") = TempStr & ", " & Trim(temparry(UBound(temparry)))



Else

    MsgBox ("Spanish Telephone Appointment Letter not needed for this type of review")

End

End If


    'Put appointment time into Calendar
    Dim olApp As Object
    
    Application.DisplayAlerts = False

    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")

    If Err.Number = 429 Then
        Set olApp = CreateObject("Outlook.application")
    End If

    On Error GoTo 0

    Set olApt = olApp.CreateItem(1)

    With olApt
        temptime = TimeValue(WorksheetFunction.Text(datasourcews.Range("AS2"), "hh:mm AM/PM"))
        tempdate = WorksheetFunction.Text(datasourcews.Range("AR4"), "mm/dd/yyyy")
        .Start = DateValue(tempdate) + TimeValue(temptime)
        .End = .Start + TimeValue("01:00:00")
        .Subject = "Telephone Appointment with " & datasourcews.Range("B2")
        .Location = "Telephone"
        .Body = "Review Number " & review_number & Chr(10) & _
            "Case Number " & Left(thisws.Range("I18"), 9)
        .BusyStatus = 2
        .ReminderMinutesBeforeStart = 5760
        .ReminderSet = True
        .SAVE
    End With

    Set olApt = Nothing
    Set olApp = Nothing
    
    datasourcewb.Close True
    Application.DisplayAlerts = True
    
    findmemo = "Telephone Appointment Letter for Review Number " & review_number & " for Sample Month " & sample_month & ".doc"

    'Copy template finding memo to active directory
    SrceFile = PathStr & "Finding Memo\Spanish Telephone CAO Appointment Letter.doc"
    tempfindmemoname = sPath & "\FM temp.doc"
    DestFile = tempfindmemoname
    FileCopy SrceFile, DestFile

    Set appWD = CreateObject("Word.Application")
    appWD.Visible = True

    With appWD
    .Documents.Open fileName:=tempfindmemoname
    
    With .ActiveDocument.MailMerge
        .OpenDataSource Name:=(tempsourcename), _
            ReadOnly:=True, LinkToSource:=0, AddToRecentFiles:=False, _
            PasswordDocument:="", PasswordTemplate:="", WritePasswordDocument:="", _
            WritePasswordTemplate:="", Revert:=False, Format:=0, _
            Connection:="yourNamedRange", SQLStatement:="SELECT * FROM `yourNamedRange`", SQLStatement1:=""
        .Destination = 0
        .SuppressBlankLines = True
        .Execute Pause:=False
    End With
    
Application.StatusBar = "Now creating document"

    .ActiveDocument.Content.Font.Name = "Times New Roman"
    '.ActiveDocument.Content.Font.Size = 12
    
    .ActiveDocument.SaveAs sPath & "\" & findmemo
    .ActiveDocument.Close
    
    .Documents(tempfindmemoname).Close SaveChanges:=False

End With
appWD.Quit
Set appWD = Nothing

'delete temp files
Kill tempfindmemoname
Kill tempsourcename

Application.StatusBar = False
Application.ScreenUpdating = True
Application.Visible = True

MsgBox "A Spanish Telephone Appointment Letter has been saved as " & sPath & "\" & findmemo & _
". Thank you!"

End Sub
Sub TeleAppt()

'Create Telephone Appointment memo
Dim appWD As Object
Dim review_number As Long
Dim sample_month As String, PathStr As String, findmemo As String
Dim countynum As String, districtnum As String

'Find path to Examiner's files on this PC
Set WshNetwork = CreateObject("WScript.Network")
Set oDrives = WshNetwork.EnumNetworkDrives

Dim DLetter As String, DUNC As String
Dim thisws As Worksheet, thiswb As Workbook
Dim sPath As String, tempsourcename As String

sPath = ActiveWorkbook.Path

Set thisws = ActiveSheet
Set thiswb = ActiveWorkbook

'picking date of appointment
SelectDate.Show
'picking time of appointment
SelectTime.Show

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

PathStr = DLetter
Application.ScreenUpdating = False

Application.StatusBar = "Starting mail merge ..."

'Copy template data source to active directory
tempsourcename = sPath & "\FM DS Temp.xlsx"
SrceFile = PathStr & "Finding Memo\Finding Memo Data Source.xlsx"
DestFile = tempsourcename
FileCopy SrceFile, DestFile

'Open spreadsheet where all the data is stored to populate findings memo
Workbooks.Open fileName:=tempsourcename, UpdateLinks:=False

Dim datasourcewb As Workbook, datasourcews As Worksheet

Set datasourcewb = ActiveWorkbook
Set datasourcews = ActiveSheet

'deleting information in Finding Memo Data Source spreadsheet
Range("B2:E2") = ""
Range("G2") = ""
Range("O2:S2") = ""
Range("W2:Y2") = ""

'Enter appointment date and time into data source file
Range("AR2") = ApptDate
Range("AS2") = ApptTime

'determing what kind of review it is by the first number in the review
review_type = Left(thisws.Name, 1)

'SNAP Positive
If review_type = 5 Then

'Client Information
Range("B2") = StrConv(thisws.Range("B4"), vbProperCase)
Range("C2") = thisws.Range("C155") & thisws.Range("D155") & "/" & Left(thisws.Range("I18"), 9)

' Get county number and name in correct format
countynum = thisws.Range("C155") & thisws.Range("D155")
districtnum = thisws.Range("E155") & thisws.Range("F155")
countylookup = countyformat(countynum, thisws.Range("M5"), districtnum)
Range("AK2") = thisws.Range("M5") & " Assistance Office"
'Look for counties with districts
'If countynum = "02" Or countynum = "23" Or countynum = "40" Or _
'        countynum = "51" Or countynum = "63" Or countynum = "65" Then
'    Range("AK2") = thisws.Range("M5") & " Office"
'Else
'    Range("AK2") = thisws.Range("M5") & " CAO"
'End If

'Put into temp data file other data
Range("D2") = thisws.Range("A18")
review_number = Range("D2")
Range("E2") = thisws.Range("AD18") & "/" & thisws.Range("AG18")
sample_month = thisws.Range("AD18") & thisws.Range("AG18")
Range("A7") = Val(thisws.Range("AJ5") & thisws.Range("AK5"))
Range("G2") = Range("C7") & "/" & thisws.Range("AE5")
Range("E7") = review_type
Range("O2") = thiswb.Sheets("FS Computation").Range("B60")
Range("P2") = thiswb.Sheets("FS Computation").Range("C60")
Range("Q2") = thisws.Range("Y22")
Range("R2") = thisws.Range("B29") & " - " & thisws.Range("G29") & " - " & thisws.Range("K29")
Range("E17") = Val(thisws.Range("K22"))
'reformat street address
If thisws.Range("B6") <> "" Then
    Range("Z2") = StrConv(thisws.Range("B5"), vbProperCase) & ", " & StrConv(thisws.Range("B6"), vbProperCase)
Else
    Range("Z2") = StrConv(thisws.Range("B5"), vbProperCase)
End If

'Reformat city and state
    temparry = Split(thisws.Range("B7"), ",")
    TempStr = StrConv(Trim(temparry(LBound(temparry))), vbProperCase)

Range("AA2") = TempStr & ", " & Trim(temparry(UBound(temparry)))



Else

    MsgBox ("Telephone Appointment Letter not needed for this type of review")

End

End If


    'Put appointment time into Calendar
    Dim olApp As Object
    
    Application.DisplayAlerts = False

    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")

    If Err.Number = 429 Then
        Set olApp = CreateObject("Outlook.application")
    End If

    On Error GoTo 0

    Set olApt = olApp.CreateItem(1)

    With olApt
        temptime = TimeValue(WorksheetFunction.Text(datasourcews.Range("AS2"), "hh:mm AM/PM"))
        tempdate = WorksheetFunction.Text(datasourcews.Range("AR4"), "mm/dd/yyyy")
        .Start = DateValue(tempdate) + TimeValue(temptime)
        .End = .Start + TimeValue("01:00:00")
        .Subject = "Telephone Appointment with " & datasourcews.Range("B2")
        .Location = "Telephone"
        .Body = "Review Number " & review_number & Chr(10) & _
            "Case Number " & Left(thisws.Range("I18"), 9)
        .BusyStatus = 2
        .ReminderMinutesBeforeStart = 5760
        .ReminderSet = True
        .SAVE
    End With

    Set olApt = Nothing
    Set olApp = Nothing
    
    datasourcewb.Close True
    Application.DisplayAlerts = True
    
    findmemo = "Telephone Appointment Letter for Review Number " & review_number & " for Sample Month " & sample_month & ".doc"

    'Copy template finding memo to active directory
    SrceFile = PathStr & "Finding Memo\Telephone Appointment Letter.doc"
    tempfindmemoname = sPath & "\FM temp.doc"
    DestFile = tempfindmemoname
    FileCopy SrceFile, DestFile

    Set appWD = CreateObject("Word.Application")
    appWD.Visible = True

    With appWD
    .Documents.Open fileName:=tempfindmemoname
    
    With .ActiveDocument.MailMerge
        .OpenDataSource Name:=(tempsourcename), _
            ReadOnly:=True, LinkToSource:=0, AddToRecentFiles:=False, _
            PasswordDocument:="", PasswordTemplate:="", WritePasswordDocument:="", _
            WritePasswordTemplate:="", Revert:=False, Format:=0, _
            Connection:="yourNamedRange", SQLStatement:="SELECT * FROM `yourNamedRange`", SQLStatement1:=""
        .Destination = 0
        .SuppressBlankLines = True
        .Execute Pause:=False
    End With
    
Application.StatusBar = "Now creating document"

    .ActiveDocument.Content.Font.Name = "Times New Roman"
    '.ActiveDocument.Content.Font.Size = 12
    
    .ActiveDocument.SaveAs sPath & "\" & findmemo
    .ActiveDocument.Close
    
    .Documents(tempfindmemoname).Close SaveChanges:=False

End With
appWD.Quit
Set appWD = Nothing

'delete temp files
Kill tempfindmemoname
Kill tempsourcename

Application.StatusBar = False
Application.ScreenUpdating = True
Application.Visible = True

MsgBox "A Telephone Appointment Letter has been saved as " & sPath & "\" & findmemo & _
". Thank you!"

End Sub
Sub QC14F()

'Create QC 14F Form
Dim appWD As Object
Dim review_number As Long
Dim sample_month As String, PathStr As String, findmemo As String
Dim countynum As String, districtnum As String

'Find path to Examiner's files on this PC
Set WshNetwork = CreateObject("WScript.Network")
Set oDrives = WshNetwork.EnumNetworkDrives

Dim DLetter As String, DUNC As String
Dim thisws As Worksheet, thiswb As Workbook
Dim sPath As String, tempsourcename As String

sPath = ActiveWorkbook.Path

Set thisws = ActiveSheet
Set thiswb = ActiveWorkbook

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

PathStr = DLetter
Application.ScreenUpdating = False

Application.StatusBar = "Starting mail merge ..."

'Copy template data source to active directory
tempsourcename = sPath & "\FM DS Temp.xlsx"
SrceFile = PathStr & "Finding Memo\Finding Memo Data Source.xlsx"
DestFile = tempsourcename
FileCopy SrceFile, DestFile

'Open spreadsheet where all the data is stored to populate findings memo
Workbooks.Open fileName:=tempsourcename, UpdateLinks:=False

Dim datasourcewb As Workbook, datasourcews As Worksheet

Set datasourcewb = ActiveWorkbook
Set datasourcews = ActiveSheet


'deleting information in Finding Memo Data Source spreadsheet
Range("B2:E2") = ""
Range("G2") = ""
Range("O2:S2") = ""
Range("W2:Y2") = ""

'determing what kind of review it is by the first number in the review
review_type = Left(thisws.Name, 1)

'SNAP Positive
If review_type = 5 Then

'Client Information
Range("B2") = StrConv(thisws.Range("B4"), vbProperCase)
Range("C2") = thisws.Range("C155") & thisws.Range("D155") & "/" & Left(thisws.Range("I18"), 9)

' Get county number and name in correct format
countynum = thisws.Range("C155") & thisws.Range("D155")
districtnum = thisws.Range("E155") & thisws.Range("F155")
countylookup = countyformat(countynum, thisws.Range("M5"), districtnum)
Range("AK2") = thisws.Range("M5") & " Assistance Office"
'Look for counties with districts
'If countynum = "02" Or countynum = "23" Or countynum = "40" Or _
'        countynum = "51" Or countynum = "63" Or countynum = "65" Then
'    Range("AK2") = thisws.Range("M5") & " Office"
'Else
'    Range("AK2") = thisws.Range("M5") & " CAO"
'End If

'Put into temp data file other data
Range("D2") = thisws.Range("A18")
review_number = Range("D2")
Range("E2") = thisws.Range("AD18") & "/" & thisws.Range("AG18")
sample_month = thisws.Range("AD18") & thisws.Range("AG18")
Range("A7") = Val(thisws.Range("AJ5") & thisws.Range("AK5"))
Range("G2") = Range("C7") & "/" & thisws.Range("AE5")
Range("E7") = review_type
Range("O2") = thiswb.Sheets("FS Computation").Range("B60")
Range("P2") = thiswb.Sheets("FS Computation").Range("C60")
Range("Q2") = thisws.Range("Y22")
Range("R2") = thisws.Range("B29") & " - " & thisws.Range("G29") & " - " & thisws.Range("K29")
Range("E17") = Val(thisws.Range("K22"))
'reformat street address
If thisws.Range("B6") <> "" Then
    Range("Z2") = StrConv(thisws.Range("B5"), vbProperCase) & ", " & StrConv(thisws.Range("B6"), vbProperCase)
Else
    Range("Z2") = StrConv(thisws.Range("B5"), vbProperCase)
End If

'Reformat city and state
    temparry = Split(thisws.Range("B7"), ",")
    TempStr = StrConv(Trim(temparry(LBound(temparry))), vbProperCase)

Range("AA2") = TempStr & ", " & Trim(temparry(UBound(temparry)))



Else

    MsgBox ("SNAP QC14F Request Letter not needed for this type of review")

End

End If

'Going out to website with list of Executive Directors
        FullTextFileName = datasourcews.Range("U6")
            
            'ActiveWorkbook.FollowHyperlink (FullTextFileName)
            Workbooks.Open fileName:=FullTextFileName, UpdateLinks:=False, ReadOnly:=True
            
            Range("A1:D1000").UnMerge

            'Allegheny Districts are formatted as dates, change them to 02/XX format
            For irow = 1 To 100
                If IsDate(Range("A" & irow)) Then
                    TempStr = WorksheetFunction.Text(Range("A" & irow), "mm/d/yyyy")
                    Range("A" & irow).NumberFormat = "@"
                    Range("A" & irow) = Left(TempStr, 4)
                End If
            Next irow
            
            'Look for county or district number
                found_row = Range("A1:A1000").Find(countylookup).Row
            
            datasourcews.Range("Y2") = Range("D" & found_row)
            'For irow = 1 To 1000
            '    If Range("A" & irow) = countylookup Then
                    'Get ED name
            '        datasourcews.Range("Y2") = Range("D" & irow)
            '        'Get CAO/DO Address
                    caoaddress = ""
                    For ii = 0 To 4
                        If Range("C" & found_row + ii) = "" Then
                            Exit For
                        End If
                        caoaddress = caoaddress & Range("C" & found_row + ii) & ", "
                    Next ii
                    datasourcews.Range("AL2") = caoaddress
                'End If
            'Next irow
            ActiveWorkbook.Close False
            
            'Application.Wait Now + TimeSerial(0, 0, 2)
            'CloseApp "CAO List", ""

            datasourcewb.Close True
            
    findmemo = "SNAP QC14F Request Letter for Review Number " & review_number & " for Sample Month " & sample_month & ".docx"

    'Copy template finding memo to active directory
    SrceFile = PathStr & "Finding Memo\SNAP QC14F.doc"
    tempfindmemoname = sPath & "\FM temp.doc"
    DestFile = tempfindmemoname
    FileCopy SrceFile, DestFile

    Set appWD = CreateObject("Word.Application")
    appWD.Visible = True

    With appWD
    .Documents.Open fileName:=tempfindmemoname
    
    With .ActiveDocument.MailMerge
        .OpenDataSource Name:=(tempsourcename), _
            ReadOnly:=True, LinkToSource:=0, AddToRecentFiles:=False, _
            PasswordDocument:="", PasswordTemplate:="", WritePasswordDocument:="", _
            WritePasswordTemplate:="", Revert:=False, Format:=0, _
            Connection:="yourNamedRange", SQLStatement:="SELECT * FROM `yourNamedRange`", SQLStatement1:=""
        .Destination = 0
        .SuppressBlankLines = True
        .Execute Pause:=False
    End With
    
    '.ActiveDocument.Protect Type:=2, NoReset:=True
    
Application.StatusBar = "Now creating document"

    .ActiveDocument.Content.Font.Name = "Times New Roman"
    '.ActiveDocument.Content.Font.Size = 12

    .ActiveDocument.SaveAs sPath & "\" & findmemo
    .ActiveDocument.Close
    
    .Documents(tempfindmemoname).Close SaveChanges:=False

End With
appWD.Quit
Set appWD = Nothing

'delete temp files
Kill tempfindmemoname
Kill tempsourcename

Application.StatusBar = False
Application.ScreenUpdating = True
Application.Visible = True

MsgBox "A SNAP QC14F Request Letter has been saved as " & sPath & "\" & findmemo & _
". Please close the Internet Explorer with the CAO Directory. - Thank you!"

End Sub

Sub QC14R()

'Create QC 14R Form
Dim appWD As Object
Dim review_number As Long
Dim sample_month As String, PathStr As String, findmemo As String
Dim countynum As String, districtnum As String

'Find path to Examiner's files on this PC
Set WshNetwork = CreateObject("WScript.Network")
Set oDrives = WshNetwork.EnumNetworkDrives

Dim DLetter As String, DUNC As String
Dim thisws As Worksheet, thiswb As Workbook
Dim sPath As String, tempsourcename As String

sPath = ActiveWorkbook.Path

Set thisws = ActiveSheet
Set thiswb = ActiveWorkbook

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

PathStr = DLetter
Application.ScreenUpdating = False

Application.StatusBar = "Starting mail merge ..."

'Copy template data source to active directory
tempsourcename = sPath & "\FM DS Temp.xlsx"
SrceFile = PathStr & "Finding Memo\Finding Memo Data Source.xlsx"
DestFile = tempsourcename
FileCopy SrceFile, DestFile

'Open spreadsheet where all the data is stored to populate findings memo
Workbooks.Open fileName:=tempsourcename, UpdateLinks:=False

Dim datasourcewb As Workbook, datasourcews As Worksheet

Set datasourcewb = ActiveWorkbook
Set datasourcews = ActiveSheet


'deleting information in Finding Memo Data Source spreadsheet
Range("B2:E2") = ""
Range("G2") = ""
Range("O2:S2") = ""
Range("W2:Y2") = ""

'determing what kind of review it is by the first number in the review
review_type = Left(thisws.Name, 1)

'SNAP Positive
If review_type = 5 Then

'Client Information
Range("B2") = StrConv(thisws.Range("B4"), vbProperCase)
Range("C2") = thisws.Range("C155") & thisws.Range("D155") & "/" & Left(thisws.Range("I18"), 9)

' Get county number and name in correct format
countynum = thisws.Range("C155") & thisws.Range("D155")
districtnum = thisws.Range("E155") & thisws.Range("F155")
countylookup = countyformat(countynum, thisws.Range("M5"), districtnum)
Range("AK2") = thisws.Range("M5") & " Assistance Office"
'Look for counties with districts
'If countynum = "02" Or countynum = "23" Or countynum = "40" Or _
'        countynum = "51" Or countynum = "63" Or countynum = "65" Then
'    Range("AK2") = thisws.Range("M5") & " Office"
'Else
'    Range("AK2") = thisws.Range("M5") & " CAO"
'End If

'Put into temp data file other data
Range("D2") = thisws.Range("A18")
review_number = Range("D2")
Range("E2") = thisws.Range("AD18") & "/" & thisws.Range("AG18")
sample_month = thisws.Range("AD18") & thisws.Range("AG18")
Range("A7") = Val(thisws.Range("AJ5") & thisws.Range("AK5"))
Range("G2") = Range("C7") & "/" & thisws.Range("AE5")
Range("E7") = review_type
Range("O2") = thiswb.Sheets("FS Computation").Range("B60")
Range("P2") = thiswb.Sheets("FS Computation").Range("C60")
Range("Q2") = thisws.Range("Y22")
Range("R2") = thisws.Range("B29") & " - " & thisws.Range("G29") & " - " & thisws.Range("K29")
Range("E17") = Val(thisws.Range("K22"))
'reformat street address
If thisws.Range("B6") <> "" Then
    Range("Z2") = StrConv(thisws.Range("B5"), vbProperCase) & ", " & StrConv(thisws.Range("B6"), vbProperCase)
Else
    Range("Z2") = StrConv(thisws.Range("B5"), vbProperCase)
End If

'Reformat city and state
    temparry = Split(thisws.Range("B7"), ",")
    TempStr = StrConv(Trim(temparry(LBound(temparry))), vbProperCase)

Range("AA2") = TempStr & ", " & Trim(temparry(UBound(temparry)))


Else

    MsgBox ("SNAP QC14R Request Letter not needed for this type of review")

End

End If

'Going out to website with list of Executive Directors
        FullTextFileName = datasourcews.Range("U6")
            
            'ActiveWorkbook.FollowHyperlink (FullTextFileName)
            Workbooks.Open fileName:=FullTextFileName, UpdateLinks:=False, ReadOnly:=True
            
            Range("A1:D1000").UnMerge

            'Allegheny Districts are formatted as dates, change them to 02/XX format
            For irow = 1 To 100
                If IsDate(Range("A" & irow)) Then
                    TempStr = WorksheetFunction.Text(Range("A" & irow), "mm/d/yyyy")
                    Range("A" & irow).NumberFormat = "@"
                    Range("A" & irow) = Left(TempStr, 4)
                End If
            Next irow
            
            'Look for county or district number
                found_row = Range("A1:A1000").Find(countylookup).Row
            
            datasourcews.Range("Y2") = Range("D" & found_row)
            'For irow = 1 To 1000
            '    If Range("A" & irow) = countylookup Then
                    'Get ED name
            '        datasourcews.Range("Y2") = Range("D" & irow)
            '        'Get CAO/DO Address
                    caoaddress = ""
                    For ii = 0 To 4
                        If Range("C" & found_row + ii) = "" Then
                            Exit For
                        End If
                        caoaddress = caoaddress & Range("C" & found_row + ii) & ", "
                    Next ii
                    datasourcews.Range("AL2") = caoaddress
                'End If
            'Next irow
            ActiveWorkbook.Close False
            
            'Application.Wait Now + TimeSerial(0, 0, 2)
            'CloseApp "CAO List", ""

            datasourcewb.Close True
            
    findmemo = "SNAP QC14R Request Letter for Review Number " & review_number & " for Sample Month " & sample_month & ".doc"

    'Copy template finding memo to active directory
    SrceFile = PathStr & "Finding Memo\SNAP QC14R.doc"
    tempfindmemoname = sPath & "\FM temp.doc"
    DestFile = tempfindmemoname
    FileCopy SrceFile, DestFile

    Set appWD = CreateObject("Word.Application")
    appWD.Visible = True

    With appWD
    .Documents.Open fileName:=tempfindmemoname
    
    With .ActiveDocument.MailMerge
        .OpenDataSource Name:=(tempsourcename), _
            ReadOnly:=True, LinkToSource:=0, AddToRecentFiles:=False, _
            PasswordDocument:="", PasswordTemplate:="", WritePasswordDocument:="", _
            WritePasswordTemplate:="", Revert:=False, Format:=0, _
            Connection:="yourNamedRange", SQLStatement:="SELECT * FROM `yourNamedRange`", SQLStatement1:=""
        .Destination = 0
        .SuppressBlankLines = True
        .Execute Pause:=False
    End With
    
Application.StatusBar = "Now creating document"

    .ActiveDocument.Content.Font.Name = "Times New Roman"
    '.ActiveDocument.Content.Font.Size = 12

    .ActiveDocument.SaveAs sPath & "\" & findmemo
    .ActiveDocument.Close
    
    .Documents(tempfindmemoname).Close SaveChanges:=False

End With
appWD.Quit
Set appWD = Nothing

'delete temp files
Kill tempfindmemoname
Kill tempsourcename

Application.StatusBar = False
Application.ScreenUpdating = True
Application.Visible = True

MsgBox "A SNAP QC14R Request Letter has been saved as " & sPath & "\" & findmemo & _
". Please close the Internet Explorer with the CAO Directory. - Thank you!"

End Sub

Sub Taxonomy()

'Create Taxonomy Info Memo
Dim appWD As Object
Dim review_number As Long, TempStr As String
Dim sample_month As String, PathStr As String, findmemo As String
Dim countynum As String, districtnum As String, county_name As String

'Find path to Examiner's files on this PC
Set WshNetwork = CreateObject("WScript.Network")
Set oDrives = WshNetwork.EnumNetworkDrives

Dim DLetter As String, DUNC As String
Dim thisws As Worksheet, thiswb As Workbook
Dim sPath As String, tempsourcename As String

sPath = ActiveWorkbook.Path

Set thisws = ActiveSheet
Set thiswb = ActiveWorkbook

'determing what kind of review it is by the first number in the review
review_type = Left(thisws.Name, 1)

'Check if Taxonomy range is blank or if there are A or C codes
num_blanks = 4
If review_type = "8" Then num_blanks = 8
If WorksheetFunction.CountBlank(thisws.Range("Taxonomy")) = num_blanks Then
    MsgBox "There are no taxonomy codes on schedule.  No memo will be produced."
    End
End If
'For Each cl In thisws.Range("Taxonomy").Cells
'    Select Case cl
'        Case "A", "a", "C", "c"
'            MsgBox ("Taxonomy Information Memo is not necessary for Letters A and C.")
'            End
'    End Select
'Next

'Find drive letter for stat network drive
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

PathStr = DLetter
Application.ScreenUpdating = False

Application.StatusBar = "Starting mail merge ..."

'Copy template data source to active directory
tempsourcename = sPath & "\FM DS Temp.xlsx"
SrceFile = PathStr & "Finding Memo\Finding Memo Data Source.xlsx"
DestFile = tempsourcename
FileCopy SrceFile, DestFile

'Open spreadsheet where all the data is stored to populate findings memo
Workbooks.Open fileName:=tempsourcename, UpdateLinks:=False

Dim datasourcewb As Workbook, datasourcews As Worksheet

Set datasourcewb = ActiveWorkbook
Set datasourcews = ActiveSheet

'deleting information in Finding Memo Data Source spreadsheet
Range("B2:E2") = ""
Range("G2") = ""
Range("O2:S2") = ""
Range("W2:Y2") = ""

'Client Information
Range("B2") = StrConv(thisws.Range("Client_Name"), vbProperCase)
Range("C2") = Right(thisws.Range("Local_Agency"), 2) & "/" & _
    Left(thisws.Range("Case_Record_Number"), 9)

' Get county number and name in correct format
countynum = Right(thisws.Range("Local_Agency"), 2)
If review_type = "2" Or review_type = "1" Or review_type = "9" Then
districtnum = thisws.Range("District_Code")
Else
districtnum = thisws.Range("District_Code1") & thisws.Range("District_Code2")
End If
'If MA Neg or SNAP neg (because no district on schedule) then lookup county name in temp file
If review_type = "8" Or review_type = "6" Then
   If countynum = "02" Or countynum = "23" Or countynum = "40" Or _
        countynum = "51" Or countynum = "63" Or countynum = "65" Then
        For jj = 50 To 139
            If Cells(jj, 1) = Val(countynum) And Cells(jj, 2) = districtnum Then
                county_name = Cells(jj, 4)
                Exit For
            End If
        Next jj
    Else
        For jj = 50 To 139
            If Cells(jj, 1) = Val(countynum) Then
                county_name = Cells(jj, 4) & " County"
                Exit For
            End If
        Next jj
    End If
Else
    county_name = thisws.Range("County")
End If

    countylookup = countyformat(countynum, county_name, districtnum)
    Range("AK2") = thisws.Range("M5") & " Assistance Office"
    'Look for counties with districts
    'If countynum = "02" Or countynum = "23" Or countynum = "40" Or _
    '    countynum = "51" Or countynum = "63" Or countynum = "65" Then
    '    Range("AK2") = county_name & " Office"
    'Else
    '    Range("AK2") = county_name
    'End If


'looking for taxonomy codes
j = 47
'For i = 3 To 6
For Each cl In thisws.Range("Taxonomy").Cells
    Select Case cl
        'Case "A", "a", "C", "c"
         '   MsgBox ("Taxonomy Information Memo is not necessary for Letters A and C.")
          '  ActiveWorkbook.Close False
           ' End
        Case "B", "b"
            Cells(2, j) = "The proper taxonomy was not followed.  Documents scanned need to be attached under the correct subject. " & _
            "Refer to http://opsweb/CAO/ImagingGuidev1.2/index.htm" & vbCrLf _
            & vbTab & "Document Title:  " & vbCrLf _
            & vbTab & "CAO Scan Date:  "
            j = j + 1
        Case "D", "d"
            Cells(2, j) = "The case record reviewed by QC consisted of paper files only.  QC found no documents scanned/imaged."
            j = j + 1
        Case "E", "e"
            Cells(2, j) = "The complete document was not scanned. The reverse side of the document(s) and/or all pages of forms were not scanned."
            j = j + 1
        Case "F", "f"
            Cells(2, j) = "Scanned documents were not titled correctly or not identified.  Please refer to the Desk Guide for Imaging Documents for proper labeling."
            j = j + 1
        Case "G", "g"
            Cells(2, j) = "Documents were not readable.  Documents were either too light or dark: documents were out of focus; tops/sides went beyond the margins and cut off words; highlighted sentences or Dollar amounts were not readable because of marker highlights."
            j = j + 1
        Case "H", "h"
            Cells(2, j) = "Multiple page documents were scanned as single page scans instead of having all pages scanned as one document."
             j = j + 1
       Case "I", "i"
            Cells(2, j) = "Documents that were scanned into the file were attached to the wrong case record."
            j = j + 1
        Case "J", "j"
            Cells(2, j) = "The files contain multiple copies of the same form, pay stub, etc. where one scan is sufficient."
            j = j + 1
         Case "K", "k"
            Cells(2, j) = "Other: "
            j = j + 1
    End Select
Next


'Put into temp data file other data
Range("D2") = thisws.Name
review_number = Range("D2")

'If MA Neg then get sample month
If review_type = "8" Or review_type = "1" Or review_type = "9" Then
    sample_month = thisws.Range("SampleMonthYear")
    Range("E2") = Left(sample_month, 2) & "/" & Right(sample_month, 4)
ElseIf review_type = "2" Then
    Range("E2") = thisws.Range("SampleMonthYear")
    sample_month = Left(thisws.Range("SampleMonthYear"), 2) & Right(thisws.Range("SampleMonthYear"), 4)
Else
    Range("E2") = thisws.Range("Sample_Month") & "/" & thisws.Range("Sample_Year")
    sample_month = thisws.Range("Sample_Month") & thisws.Range("Sample_Year")
End If

If review_type = "6" Then
    Range("A7") = Val(thisws.Range("Examiner_Number"))
Else
    Range("A7") = Val(thisws.Range("Examiner_Number1") & thisws.Range("Examiner_Number2"))
End If
    Range("E7") = review_type

If review_type = "5" Then
    Range("O2") = thiswb.Sheets("FS Computation").Range("B60")
    Range("P2") = thiswb.Sheets("FS Computation").Range("C60")
End If

'reformat street address
If review_type <> "1" And review_type <> "6" Then
    If thisws.Range("Client_2ndAddress") <> "" Then
        Range("Z2") = StrConv(thisws.Range("Client_1stAddress"), vbProperCase) & ", " & StrConv(thisws.Range("Client_2ndAddress"), vbProperCase)
    Else
        Range("Z2") = StrConv(thisws.Range("Client_1stAddress"), vbProperCase)
    End If
Else
     Range("Z2") = StrConv(thisws.Range("Client_1stAddress"), vbProperCase)
End If
'Reformat city and state
    temparry = Split(thisws.Range("Client_TownStateZip"), ",")
    TempStr = StrConv(Trim(temparry(LBound(temparry))), vbProperCase)

Range("AA2") = TempStr & ", " & Trim(temparry(UBound(temparry)))


'Going out to website with list of Executive Directors
        FullTextFileName = datasourcews.Range("U6")
            
            Workbooks.Open fileName:=FullTextFileName, UpdateLinks:=False, ReadOnly:=True
            
            Range("A1:D1000").UnMerge

            'Allegheny Districts are formatted as dates, change them to 02/XX format
            For irow = 1 To 100
                If IsDate(Range("A" & irow)) Then
                    TempStr = WorksheetFunction.Text(Range("A" & irow), "mm/d/yyyy")
                    Range("A" & irow).NumberFormat = "@"
                    Range("A" & irow) = Left(TempStr, 4)
                End If
            Next irow
            
            'Look for county or district number
                found_row = Range("A1:A1000").Find(countylookup).Row
            
            datasourcews.Range("Y2") = Range("D" & found_row)
            'For irow = 1 To 1000
            '    If Range("A" & irow) = countylookup Then
                    'Get ED name
            '        datasourcews.Range("Y2") = Range("D" & irow)
            '        'Get CAO/DO Address
                    caoaddress = ""
                    For ii = 0 To 4
                        If Range("C" & found_row + ii) = "" Then
                            Exit For
                        End If
                        caoaddress = caoaddress & Range("C" & found_row + ii) & ", "
                    Next ii
                    datasourcews.Range("AL2") = caoaddress
                'End If
            'Next irow
            ActiveWorkbook.Close False
            
    
    datasourcewb.Close True
    Application.DisplayAlerts = True
            
    findmemo = "Taxonomy Information Memo for Review Number " & review_number & " for Sample Month " & sample_month & ".doc"

    'Copy template cao appointment memo to active directory
    SrceFile = PathStr & "Finding Memo\Taxonomy Information Memo Master.doc"
    tempfindmemoname = sPath & "\FM temp.doc"
    DestFile = tempfindmemoname
    FileCopy SrceFile, DestFile

    Set appWD = CreateObject("Word.Application")
    appWD.Visible = True

    With appWD
    .Documents.Open fileName:=tempfindmemoname
    
    'Merge CAO information into template
    With .ActiveDocument.MailMerge
        .OpenDataSource Name:=(tempsourcename), _
            ReadOnly:=True, LinkToSource:=0, AddToRecentFiles:=False, _
            PasswordDocument:="", PasswordTemplate:="", WritePasswordDocument:="", _
            WritePasswordTemplate:="", Revert:=False, Format:=0, _
            Connection:="yourNamedRange", SQLStatement:="SELECT * FROM `yourNamedRange`", SQLStatement1:=""
        .Destination = 0
        .SuppressBlankLines = True
        .Execute Pause:=False
    End With
    
Application.StatusBar = "Now creating document"

    .ActiveDocument.Content.Font.Name = "Times New Roman"
    '.ActiveDocument.Content.Font.Size = 12

    .ActiveDocument.SaveAs sPath & "\" & findmemo
    .ActiveDocument.Close
    
    .Documents(tempfindmemoname).Close SaveChanges:=False

End With
appWD.Quit
Set appWD = Nothing

'delete temp files
Kill tempfindmemoname
Kill tempsourcename

Application.StatusBar = False
Application.ScreenUpdating = True
Application.Visible = True

MsgBox "A Taxonomy Information Memo has been saved as " & sPath & "\" & findmemo & _
"."

End Sub

Sub MAAppt()

'Create CAO Appointment Memo
Dim appWD As Object
Dim review_number As Long, TempStr As String
Dim sample_month As String, PathStr As String, findmemo As String
Dim countynum As String, districtnum As String

'Find path to Examiner's files on this PC
Set WshNetwork = CreateObject("WScript.Network")
Set oDrives = WshNetwork.EnumNetworkDrives

Dim DLetter As String, DUNC As String
Dim thisws As Worksheet, thiswb As Workbook
Dim sPath As String, tempsourcename As String

sPath = ActiveWorkbook.Path

Set thisws = ActiveSheet
Set thiswb = ActiveWorkbook

'picking date of appointment
SelectDate.Show
'picking time of appointment
SelectTime.Show

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

PathStr = DLetter
Application.ScreenUpdating = False

Application.StatusBar = "Starting mail merge ..."

'Copy template data source to active directory
tempsourcename = sPath & "\FM DS Temp.xlsx"
SrceFile = PathStr & "Finding Memo\Finding Memo Data Source.xlsx"
DestFile = tempsourcename
FileCopy SrceFile, DestFile

'Open spreadsheet where all the data is stored to populate findings memo
Workbooks.Open fileName:=tempsourcename, UpdateLinks:=False

Dim datasourcewb As Workbook, datasourcews As Worksheet

Set datasourcewb = ActiveWorkbook
Set datasourcews = ActiveSheet

'deleting information in Finding Memo Data Source spreadsheet
Range("B2:E2") = ""
Range("G2") = ""
Range("O2:S2") = ""
Range("W2:Y2") = ""

'Enter appointment date and time into data source file
Range("AR2") = ApptDate
Range("AS2") = ApptTime

'determing what kind of review it is by the first number in the review
review_type = Left(thisws.Name, 1)

'MA Positive
If review_type = 2 Then

'Client Information
Range("B2") = StrConv(thisws.Range("B2"), vbProperCase)
'County and Case Number
Range("C2") = Right(thisws.Range("AC10"), 2) & "/" & thisws.Range("H10")

' Get county number and name in correct format
countynum = Right(thisws.Range("AC10"), 2)
districtnum = thisws.Range("AH10")
countylookup = countyformat(countynum, thisws.Range("O4"), districtnum)
Range("AK2") = thisws.Range("M5") & " Assistance Office"
'Look for counties with districts
'If countynum = "02" Or countynum = "23" Or countynum = "40" Or _
'        countynum = "51" Or countynum = "63" Or countynum = "65" Then
'    Range("AK2") = thisws.Range("O4") & " Office"
'Else
'    Range("AK2") = thisws.Range("O4") & " CAO"
'End If

'review number
Range("D2") = thisws.Range("A10") & thisws.Range("B10") & thisws.Range("c10") & thisws.Range("D10") & thisws.Range("E10") & thisws.Range("F10")
review_number = Range("D2")
'review month
Range("E2") = thisws.Range("AM10")
sample_month = Range("E2")
'reviewer number
Range("A7") = Val(thisws.Range("AO3") & thisws.Range("AP3"))
'supervisor/examiner
Range("G2") = Range("C7") & "/" & thisws.Range("AF2")
Range("E7") = review_type
'
Range("R2") = thisws.Range("J61") & " - " & thisws.Range("O61")
'finding
Range("E17") = Val(thisws.Range("AL10"))
'reformat street address
If thisws.Range("B4") <> "" Then
    Range("Z2") = StrConv(thisws.Range("B3"), vbProperCase) & ", " & StrConv(thisws.Range("B4"), vbProperCase)
Else
    Range("Z2") = StrConv(thisws.Range("B3"), vbProperCase)
End If

'Reformat city and state
    temparry = Split(thisws.Range("B5"), ",")
    TempStr = StrConv(Trim(temparry(LBound(temparry))), vbProperCase)

Range("AA2") = TempStr & ", " & Trim(temparry(UBound(temparry)))

Else

    MsgBox ("MA Appointment Letter not needed for this type of review")

End

End If

    'Put appointment time into Calendar
    Dim olApp As Object
    
    Application.DisplayAlerts = False

    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")

    If Err.Number = 429 Then
        Set olApp = CreateObject("Outlook.application")
    End If

    On Error GoTo 0

    Set olApt = olApp.CreateItem(1)

    With olApt
        temptime = TimeValue(WorksheetFunction.Text(datasourcews.Range("AS2"), "hh:mm AM/PM"))
        tempdate = WorksheetFunction.Text(datasourcews.Range("AR4"), "mm/dd/yyyy")
        .Start = DateValue(tempdate) + TimeValue(temptime)
        .End = .Start + TimeValue("01:00:00")
        .Subject = "Appointment with " & datasourcews.Range("B2")
        .Location = datasourcews.Range("AK2")
        .Body = "Review Number " & review_number & Chr(10) & _
            "Case Number " & Left(thisws.Range("I18"), 9)
        .BusyStatus = 3
        .ReminderMinutesBeforeStart = 5760
        .ReminderSet = True
        .SAVE
    End With

    Set olApt = Nothing
    Set olApp = Nothing
    
    datasourcewb.Close True
    Application.DisplayAlerts = True
            
    findmemo = "MA Appointment Letter for Review Number " & _
        review_number & " for Sample Month " & _
        WorksheetFunction.Text(sample_month, "yyyymm") & ".doc"

    'Copy template cao appointment memo to active directory
    SrceFile = PathStr & "Finding Memo\MA Appointment Letter Master.doc"
    tempfindmemoname = sPath & "\FM temp.doc"
    DestFile = tempfindmemoname
    FileCopy SrceFile, DestFile

    Set appWD = CreateObject("Word.Application")
    appWD.Visible = True

    With appWD
    .Documents.Open fileName:=tempfindmemoname
    
    'Merge CAO information into template
    With .ActiveDocument.MailMerge
        .OpenDataSource Name:=(tempsourcename), _
            ReadOnly:=True, LinkToSource:=0, AddToRecentFiles:=False, _
            PasswordDocument:="", PasswordTemplate:="", WritePasswordDocument:="", _
            WritePasswordTemplate:="", Revert:=False, Format:=0, _
            Connection:="yourNamedRange", SQLStatement:="SELECT * FROM `yourNamedRange`", SQLStatement1:=""
        .Destination = 0
        .SuppressBlankLines = True
        .Execute Pause:=False
    End With
    
Application.StatusBar = "Now creating document"

    .ActiveDocument.Content.Font.Name = "Times New Roman"
    '.ActiveDocument.Content.Font.Size = 12

    .ActiveDocument.SaveAs sPath & "\" & findmemo
    .ActiveDocument.Close
    
    .Documents(tempfindmemoname).Close SaveChanges:=False

End With
appWD.Quit
Set appWD = Nothing

'delete temp files
Kill tempfindmemoname
Kill tempsourcename

Application.StatusBar = False
Application.ScreenUpdating = True
Application.Visible = True

MsgBox "A MA Appointment Letter has been saved as " & sPath & "\" & findmemo & _
". Please close the Internet Explorer with the CAO Directory. - Thank you!"

End Sub
Sub MASpCAOAppt()

'Create CAO Appointment Memo
Dim appWD As Object
Dim review_number As Long, TempStr As String
Dim sample_month As String, PathStr As String, findmemo As String
Dim countynum As String, districtnum As String

'Find path to Examiner's files on this PC
Set WshNetwork = CreateObject("WScript.Network")
Set oDrives = WshNetwork.EnumNetworkDrives

Dim DLetter As String, DUNC As String
Dim thisws As Worksheet, thiswb As Workbook
Dim sPath As String, tempsourcename As String

sPath = ActiveWorkbook.Path

Set thisws = ActiveSheet
Set thiswb = ActiveWorkbook

'picking date of appointment
SelectDate.Show
'picking time of appointment
SelectTime.Show

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

PathStr = DLetter
Application.ScreenUpdating = False

Application.StatusBar = "Starting mail merge ..."

'Copy template data source to active directory
tempsourcename = sPath & "\FM DS Temp.xlsx"
SrceFile = PathStr & "Finding Memo\Finding Memo Data Source.xlsx"
DestFile = tempsourcename
FileCopy SrceFile, DestFile

'Open spreadsheet where all the data is stored to populate findings memo
Workbooks.Open fileName:=tempsourcename, UpdateLinks:=False

Dim datasourcewb As Workbook, datasourcews As Worksheet

Set datasourcewb = ActiveWorkbook
Set datasourcews = ActiveSheet

'deleting information in Finding Memo Data Source spreadsheet
Range("B2:E2") = ""
Range("G2") = ""
Range("O2:S2") = ""
Range("W2:Y2") = ""

'Enter appointment date and time into data source file
Range("AR2") = ApptDate
Range("AS2") = ApptTime

'determing what kind of review it is by the first number in the review
review_type = Left(thisws.Name, 1)

'MA Positive
If review_type = 2 Then

'Client Information
Range("B2") = StrConv(thisws.Range("B2"), vbProperCase)
'County and Case Number
Range("C2") = Right(thisws.Range("AC10"), 2) & "/" & thisws.Range("H10")

' Get county number and name in correct format
countynum = Right(thisws.Range("AC10"), 2)
districtnum = thisws.Range("AH10")
countylookup = countyformat(countynum, thisws.Range("O4"), districtnum)
Range("AK2") = thisws.Range("M5") & " Assistance Office"
'Look for counties with districts
'If countynum = "02" Or countynum = "23" Or countynum = "40" Or _
'        countynum = "51" Or countynum = "63" Or countynum = "65" Then
'    Range("AK2") = thisws.Range("O4") & " Office"
'Else
'    Range("AK2") = thisws.Range("O4") & " CAO"
'End If

'review number
Range("D2") = thisws.Range("A10") & thisws.Range("B10") & thisws.Range("c10") & thisws.Range("D10") & thisws.Range("E10") & thisws.Range("F10")
review_number = Range("D2")
'review month
Range("E2") = thisws.Range("AM10")
sample_month = Range("E2")
'reviewer number
Range("A7") = Val(thisws.Range("AO3") & thisws.Range("AP3"))
'supervisor/examiner
Range("G2") = Range("C7") & "/" & thisws.Range("AF2")
Range("E7") = review_type
'
Range("R2") = thisws.Range("J61") & " - " & thisws.Range("O61")
'finding
Range("E17") = Val(thisws.Range("AL10"))
'reformat street address
If thisws.Range("B4") <> "" Then
    Range("Z2") = StrConv(thisws.Range("B3"), vbProperCase) & ", " & StrConv(thisws.Range("B4"), vbProperCase)
Else
    Range("Z2") = StrConv(thisws.Range("B3"), vbProperCase)
End If

'Reformat city and state
    temparry = Split(thisws.Range("B5"), ",")
    TempStr = StrConv(Trim(temparry(LBound(temparry))), vbProperCase)

Range("AA2") = TempStr & ", " & Trim(temparry(UBound(temparry)))

Else

    MsgBox ("MA Appointment Letter not needed for this type of review")

End

End If

    'Put appointment time into Calendar
    Dim olApp As Object
    
    Application.DisplayAlerts = False

    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")

    If Err.Number = 429 Then
        Set olApp = CreateObject("Outlook.application")
    End If

    On Error GoTo 0

    Set olApt = olApp.CreateItem(1)

    With olApt
        temptime = TimeValue(WorksheetFunction.Text(datasourcews.Range("AS2"), "hh:mm AM/PM"))
        tempdate = WorksheetFunction.Text(datasourcews.Range("AR4"), "mm/dd/yyyy")
        .Start = DateValue(tempdate) + TimeValue(temptime)
        .End = .Start + TimeValue("01:00:00")
        .Subject = "Appointment with " & datasourcews.Range("B2")
        .Location = datasourcews.Range("AK2")
        .Body = "Review Number " & review_number & Chr(10) & _
            "Case Number " & Left(thisws.Range("I18"), 9)
        .BusyStatus = 3
        .ReminderMinutesBeforeStart = 5760
        .ReminderSet = True
        .SAVE
    End With

    Set olApt = Nothing
    Set olApp = Nothing
    
    datasourcewb.Close True
    Application.DisplayAlerts = True
            
    findmemo = "MA Appointment Letter for Review Number " & _
        review_number & " for Sample Month " & _
        WorksheetFunction.Text(sample_month, "yyyymm") & ".doc"

    'Copy template cao appointment memo to active directory
    SrceFile = PathStr & "Finding Memo\MA Appointment Letter Spanish.doc"
    tempfindmemoname = sPath & "\FM temp.doc"
    DestFile = tempfindmemoname
    FileCopy SrceFile, DestFile

    Set appWD = CreateObject("Word.Application")
    appWD.Visible = True

    With appWD
    .Documents.Open fileName:=tempfindmemoname
    
    'Merge CAO information into template
    With .ActiveDocument.MailMerge
        .OpenDataSource Name:=(tempsourcename), _
            ReadOnly:=True, LinkToSource:=0, AddToRecentFiles:=False, _
            PasswordDocument:="", PasswordTemplate:="", WritePasswordDocument:="", _
            WritePasswordTemplate:="", Revert:=False, Format:=0, _
            Connection:="yourNamedRange", SQLStatement:="SELECT * FROM `yourNamedRange`", SQLStatement1:=""
        .Destination = 0
        .SuppressBlankLines = True
        .Execute Pause:=False
    End With
    
Application.StatusBar = "Now creating document"

    .ActiveDocument.Content.Font.Name = "Times New Roman"
    '.ActiveDocument.Content.Font.Size = 12

    .ActiveDocument.SaveAs sPath & "\" & findmemo
    .ActiveDocument.Close
    
    .Documents(tempfindmemoname).Close SaveChanges:=False

End With
appWD.Quit
Set appWD = Nothing

'delete temp files
Kill tempfindmemoname
Kill tempsourcename

Application.StatusBar = False
Application.ScreenUpdating = True
Application.Visible = True

MsgBox "A MA Appointment Letter has been saved as " & sPath & "\" & findmemo & _
". Please close the Internet Explorer with the CAO Directory. - Thank you!"

End Sub

Sub NHLLR()

'Create Nursing Home LRR
Dim appWD As Object
Dim review_number As Long, TempStr As String
Dim sample_month As String, PathStr As String, findmemo As String
Dim countynum As String, districtnum As String

'Find path to Examiner's files on this PC
Set WshNetwork = CreateObject("WScript.Network")
Set oDrives = WshNetwork.EnumNetworkDrives

Dim DLetter As String, DUNC As String
Dim thisws As Worksheet, thiswb As Workbook
Dim sPath As String, tempsourcename As String

sPath = ActiveWorkbook.Path

Set thisws = ActiveSheet
Set thiswb = ActiveWorkbook


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

PathStr = DLetter
Application.ScreenUpdating = False

Application.StatusBar = "Starting mail merge ..."

'Copy template data source to active directory
tempsourcename = sPath & "\FM DS Temp.xlsx"
SrceFile = PathStr & "Finding Memo\Finding Memo Data Source.xlsx"
DestFile = tempsourcename
FileCopy SrceFile, DestFile

'Open spreadsheet where all the data is stored to populate findings memo
Workbooks.Open fileName:=tempsourcename, UpdateLinks:=False

Dim datasourcewb As Workbook, datasourcews As Worksheet

Set datasourcewb = ActiveWorkbook
Set datasourcews = ActiveSheet

'deleting information in Finding Memo Data Source spreadsheet
Range("B2:E2") = ""
Range("G2") = ""
Range("O2:S2") = ""
Range("W2:Y2") = ""

'Enter appointment date and time into data source file
Range("AR2") = ApptDate
Range("AS2") = ApptTime

'determing what kind of review it is by the first number in the review
review_type = Left(thisws.Name, 1)

'MA Positive
If review_type = 2 Then

'Client Information
Range("B2") = StrConv(thisws.Range("B2"), vbProperCase)
'County and Case Number
Range("C2") = Right(thisws.Range("AC10"), 2) & "/" & thisws.Range("H10")

' Get county number and name in correct format
countynum = Right(thisws.Range("AC10"), 2)
districtnum = thisws.Range("AH10")
countylookup = countyformat(countynum, thisws.Range("O4"), districtnum)
Range("AK2") = thisws.Range("M5") & " Assistance Office"


'review number
Range("D2") = thisws.Range("A10") & thisws.Range("B10") & thisws.Range("c10") & thisws.Range("D10") & thisws.Range("E10") & thisws.Range("F10")
review_number = Range("D2")
'review month
Range("E2") = thisws.Range("AM10")
sample_month = Range("E2")
'reviewer number
Range("A7") = Val(thisws.Range("AO3") & thisws.Range("AP3"))
'supervisor/examiner
Range("G2") = Range("C7") & "/" & thisws.Range("AF2")
Range("E7") = review_type
Range("R2") = thisws.Range("J61") & " - " & thisws.Range("O61")
'finding
Range("E17") = Val(thisws.Range("AL10"))
'reformat street address
If thisws.Range("B4") <> "" Then
    Range("Z2") = StrConv(thisws.Range("B3"), vbProperCase) & ", " & StrConv(thisws.Range("B4"), vbProperCase)
Else
    Range("Z2") = StrConv(thisws.Range("B3"), vbProperCase)
End If

'Reformat city and state
    temparry = Split(thisws.Range("B5"), ",")
    TempStr = StrConv(Trim(temparry(LBound(temparry))), vbProperCase)

Range("AA2") = TempStr & ", " & Trim(temparry(UBound(temparry)))

Else

    MsgBox ("MA Nursing Home LRR Letter is not needed for this type of review")

End

End If

    datasourcewb.Close True
    Application.DisplayAlerts = True
    
    findmemo = "Nursing Home LRR for Review Number " & _
        review_number & " for Sample Month " & _
        WorksheetFunction.Text(sample_month, "yyyymm") & ".doc"

    'Copy template cao appointment memo to active directory
    SrceFile = PathStr & "Finding Memo\Nursing Home LRR.docx"
    tempfindmemoname = sPath & "\FM temp.doc"
    DestFile = tempfindmemoname
    FileCopy SrceFile, DestFile

    Set appWD = CreateObject("Word.Application")
    appWD.Visible = True

    With appWD
    .Documents.Open fileName:=tempfindmemoname
    
    'Merge CAO information into template
    With .ActiveDocument.MailMerge
        .OpenDataSource Name:=(tempsourcename), _
            ReadOnly:=True, LinkToSource:=0, AddToRecentFiles:=False, _
            PasswordDocument:="", PasswordTemplate:="", WritePasswordDocument:="", _
            WritePasswordTemplate:="", Revert:=False, Format:=0, _
            Connection:="yourNamedRange", SQLStatement:="SELECT * FROM `yourNamedRange`", SQLStatement1:=""
        .Destination = 0
        .SuppressBlankLines = True
        .Execute Pause:=False
    End With
    
Application.StatusBar = "Now creating document"

    .ActiveDocument.Content.Font.Name = "Times New Roman"
    '.ActiveDocument.Content.Font.Size = 12

    .ActiveDocument.SaveAs sPath & "\" & findmemo
    .ActiveDocument.Close
    
    .Documents(tempfindmemoname).Close SaveChanges:=False

End With
appWD.Quit
Set appWD = Nothing

'delete temp files
Kill tempfindmemoname
Kill tempsourcename

Application.StatusBar = False
Application.ScreenUpdating = True
Application.Visible = True

MsgBox "A Nursing Home LRR letter has been saved as " & sPath & "\" & findmemo & _
". - Thank you!"

End Sub
Sub NHBUS()

'Create Nursing Home Business Form
Dim appWD As Object
Dim review_number As Long, TempStr As String
Dim sample_month As String, PathStr As String, findmemo As String
Dim countynum As String, districtnum As String

'Find path to Examiner's files on this PC
Set WshNetwork = CreateObject("WScript.Network")
Set oDrives = WshNetwork.EnumNetworkDrives

Dim DLetter As String, DUNC As String
Dim thisws As Worksheet, thiswb As Workbook
Dim sPath As String, tempsourcename As String

sPath = ActiveWorkbook.Path

Set thisws = ActiveSheet
Set thiswb = ActiveWorkbook



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

PathStr = DLetter
Application.ScreenUpdating = False

Application.StatusBar = "Starting mail merge ..."

'Copy template data source to active directory
tempsourcename = sPath & "\FM DS Temp.xlsx"
SrceFile = PathStr & "Finding Memo\Finding Memo Data Source.xlsx"
DestFile = tempsourcename
FileCopy SrceFile, DestFile

'Open spreadsheet where all the data is stored to populate findings memo
Workbooks.Open fileName:=tempsourcename, UpdateLinks:=False

Dim datasourcewb As Workbook, datasourcews As Worksheet

Set datasourcewb = ActiveWorkbook
Set datasourcews = ActiveSheet

'deleting information in Finding Memo Data Source spreadsheet
Range("B2:E2") = ""
Range("G2") = ""
Range("O2:S2") = ""
Range("W2:Y2") = ""

'Enter appointment date and time into data source file
Range("AR2") = ApptDate
Range("AS2") = ApptTime

'determing what kind of review it is by the first number in the review
review_type = Left(thisws.Name, 1)

'MA Positive
If review_type = 2 Then

'Client Information
Range("B2") = StrConv(thisws.Range("B2"), vbProperCase)
'County and Case Number
Range("C2") = Right(thisws.Range("AC10"), 2) & "/" & thisws.Range("H10")

' Get county number and name in correct format
countynum = Right(thisws.Range("AC10"), 2)
districtnum = thisws.Range("AH10")
countylookup = countyformat(countynum, thisws.Range("O4"), districtnum)
Range("AK2") = thisws.Range("M5") & " Assistance Office"


'review number
Range("D2") = thisws.Range("A10") & thisws.Range("B10") & thisws.Range("c10") & thisws.Range("D10") & thisws.Range("E10") & thisws.Range("F10")
review_number = Range("D2")
'review month
Range("E2") = thisws.Range("AM10")
sample_month = Range("E2")
'reviewer number
Range("A7") = Val(thisws.Range("AO3") & thisws.Range("AP3"))
'supervisor/examiner
Range("G2") = Range("C7") & "/" & thisws.Range("AF2")
Range("E7") = review_type
Range("R2") = thisws.Range("J61") & " - " & thisws.Range("O61")
'finding
Range("E17") = Val(thisws.Range("AL10"))
'reformat street address
If thisws.Range("B4") <> "" Then
    Range("Z2") = StrConv(thisws.Range("B3"), vbProperCase) & ", " & StrConv(thisws.Range("B4"), vbProperCase)
Else
    Range("Z2") = StrConv(thisws.Range("B3"), vbProperCase)
End If

'Reformat city and state
    temparry = Split(thisws.Range("B5"), ",")
    TempStr = StrConv(Trim(temparry(LBound(temparry))), vbProperCase)

Range("AA2") = TempStr & ", " & Trim(temparry(UBound(temparry)))

Else

    MsgBox ("MA Nursing Home Business Office Letter is not needed for this type of review")

End

End If

    datasourcewb.Close True
    Application.DisplayAlerts = True
    
    findmemo = "Nursing Home Business Office for Review Number " & _
        review_number & " for Sample Month " & _
        WorksheetFunction.Text(sample_month, "yyyymm") & ".doc"

    'Copy template cao appointment memo to active directory
    SrceFile = PathStr & "Finding Memo\Nursing Home Business Office.docx"
    tempfindmemoname = sPath & "\FM temp.doc"
    DestFile = tempfindmemoname
    FileCopy SrceFile, DestFile

    Set appWD = CreateObject("Word.Application")
    appWD.Visible = True

    With appWD
    .Documents.Open fileName:=tempfindmemoname
    
    'Merge CAO information into template
    With .ActiveDocument.MailMerge
        .OpenDataSource Name:=(tempsourcename), _
            ReadOnly:=True, LinkToSource:=0, AddToRecentFiles:=False, _
            PasswordDocument:="", PasswordTemplate:="", WritePasswordDocument:="", _
            WritePasswordTemplate:="", Revert:=False, Format:=0, _
            Connection:="yourNamedRange", SQLStatement:="SELECT * FROM `yourNamedRange`", SQLStatement1:=""
        .Destination = 0
        .SuppressBlankLines = True
        .Execute Pause:=False
    End With
    
Application.StatusBar = "Now creating document"

    .ActiveDocument.Content.Font.Name = "Times New Roman"
    '.ActiveDocument.Content.Font.Size = 12

    .ActiveDocument.SaveAs sPath & "\" & findmemo
    .ActiveDocument.Close
    
    .Documents(tempfindmemoname).Close SaveChanges:=False

End With
appWD.Quit
Set appWD = Nothing

'delete temp files
Kill tempfindmemoname
Kill tempsourcename

Application.StatusBar = False
Application.ScreenUpdating = True
Application.Visible = True

MsgBox "A Nursing Home Business Office letter has been saved as " & sPath & "\" & findmemo & _
". - Thank you!"

End Sub

Sub Rush()

'Create Case Summary Form
Dim appWD As Object
Dim review_number As Long
Dim sample_month As String, PathStr As String, findmemo As String
Dim countynum As String, districtnum As String

'Find path to Examiner's files on this PC
Set WshNetwork = CreateObject("WScript.Network")
Set oDrives = WshNetwork.EnumNetworkDrives

Dim DLetter As String, DUNC As String
Dim thisws As Worksheet, thiswb As Workbook
Dim sPath As String, tempsourcename As String

sPath = ActiveWorkbook.Path

Set thisws = ActiveSheet
Set thiswb = ActiveWorkbook

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

PathStr = DLetter
Application.ScreenUpdating = False

Application.StatusBar = "Starting mail merge ..."

'Copy template data source to active directory
tempsourcename = sPath & "\FM DS Temp.xlsx"
SrceFile = PathStr & "Finding Memo\Finding Memo Data Source.xlsx"
DestFile = tempsourcename
FileCopy SrceFile, DestFile

'Open spreadsheet where all the data is stored to populate findings memo
Workbooks.Open fileName:=tempsourcename, UpdateLinks:=False

Dim datasourcewb As Workbook, datasourcews As Worksheet

Set datasourcewb = ActiveWorkbook
Set datasourcews = ActiveSheet


'deleting information in Finding Memo Data Source spreadsheet
Range("B2:E2") = ""
Range("G2") = ""
Range("O2:S2") = ""
Range("W2:Y2") = ""

'determing what kind of review it is by the first number in the review
review_type = Left(thisws.Name, 1)

'SNAP Positive
If review_type = 5 Then

'Client Information
Range("B2") = StrConv(thisws.Range("B4"), vbProperCase)
Range("C2") = thisws.Range("C155") & thisws.Range("D155") & "/" & Left(thisws.Range("I18"), 9)

' Get county number and name in correct format
countynum = thisws.Range("C155") & thisws.Range("D155")
districtnum = thisws.Range("E155") & thisws.Range("F155")
countylookup = countyformat(countynum, thisws.Range("M5"), districtnum)
Range("AK2") = thisws.Range("M5") & " Assistance Office"
'Look for counties with districts
'If countynum = "02" Or countynum = "23" Or countynum = "40" Or _
'        countynum = "51" Or countynum = "63" Or countynum = "65" Then
'    Range("AK2") = thisws.Range("M5") & " Office"
'Else
'    Range("AK2") = thisws.Range("M5") & " CAO"
'End If

'Put into temp data file other data
Range("D2") = thisws.Range("A18")
review_number = Range("D2")
Range("E2") = thisws.Range("AD18") & "/" & thisws.Range("AG18")
sample_month = thisws.Range("AD18") & thisws.Range("AG18")
Range("A7") = Val(thisws.Range("AJ5") & thisws.Range("AK5"))
Range("G2") = Range("C7") & "/" & thisws.Range("AE5")
Range("E7") = review_type
Range("O2") = thiswb.Sheets("FS Computation").Range("B60")
Range("P2") = thiswb.Sheets("FS Computation").Range("C60")
Range("Q2") = thisws.Range("Y22")
Range("R2") = thisws.Range("B29") & " - " & thisws.Range("G29") & " - " & thisws.Range("K29")
Range("E17") = Val(thisws.Range("K22"))
'reformat street address
If thisws.Range("B6") <> "" Then
    Range("Z2") = StrConv(thisws.Range("B5"), vbProperCase) & ", " & StrConv(thisws.Range("B6"), vbProperCase)
Else
    Range("Z2") = StrConv(thisws.Range("B5"), vbProperCase)
End If

'Reformat city and state
    temparry = Split(thisws.Range("B7"), ",")
    TempStr = StrConv(Trim(temparry(LBound(temparry))), vbProperCase)

Range("AA2") = TempStr & ", " & Trim(temparry(UBound(temparry)))


'SNAP Negative
ElseIf review_type = 6 Then

'Client Information
Range("B2") = StrConv(thisws.Range("C8"), vbProperCase)
Range("C2") = Right(thisws.Range("AA20"), 2) & "/" & Right(thisws.Range("L20"), 8)

' Get county number and name in correct format
countynum = Right(thisws.Range("AA20"), 2)
districtnum = thisws.Range("AD20") & thisws.Range("AE20")
countylookup = countyformat(countynum, thisws.Range("C1"), districtnum)
Range("AK2") = thisws.Range("M5") & " Assistance Office"
'Look for counties with districts
'If countynum = "02" Or countynum = "23" Or countynum = "40" Or _
'        countynum = "51" Or countynum = "63" Or countynum = "65" Then
'    Range("AK2") = thisws.Range("C1") & " Office"
'Else
'    Range("AK2") = thisws.Range("C1") & " CAO"
'End If

'Put into temp data file other data
Range("D2") = thisws.Range("C20")
review_number = Range("D2")
Range("E2") = thisws.Range("AF20") & "/" & thisws.Range("AI20")
sample_month = thisws.Range("AF20") & thisws.Range("AI20")
Range("A7") = Val(thisws.Range("W17"))
Range("G2") = Range("C7") & "/" & thisws.Range("Q17")
Range("E7") = review_type
Range("F7") = review_type
Range("R2") = thisws.Range("E41") & " - " & thisws.Range("W41")

'reformat street address
    Range("Z2") = StrConv(thisws.Range("C12"), vbProperCase)

'Reformat city and state
    temparry = Split(thisws.Range("C13"), ",")
    TempStr = StrConv(Trim(temparry(LBound(temparry))), vbProperCase)

Range("AA2") = TempStr & ", " & Trim(temparry(UBound(temparry)))

        
'determining if there is an error on the schedule
If thisws.Range("AF29") = 2 And (thisws.Range("H34") = 3 Or thisws.Range("H34") = "") Then
    If Mid(thiswb.Name, 15, 2) = "66" Or Mid(thiswb.Name, 15, 2) = "65" Then
        Range("W2") = "X"
    Else
        Range("X2") = "X"
    End If
End If


End If

    datasourcewb.Close True
    findmemo = "Case Summary Letter for Review Number " & review_number & " for Sample Month " & sample_month & ".doc"

    'Copy template finding memo to active directory
    SrceFile = PathStr & "Finding Memo\Case Summary Template.doc"
    tempfindmemoname = sPath & "\FM temp.doc"
    DestFile = tempfindmemoname
    FileCopy SrceFile, DestFile

    Set appWD = CreateObject("Word.Application")
    appWD.Visible = True

    With appWD
    .Documents.Open fileName:=tempfindmemoname
    
    With .ActiveDocument.MailMerge
        .OpenDataSource Name:=(tempsourcename), _
            ReadOnly:=True, LinkToSource:=0, AddToRecentFiles:=False, _
            PasswordDocument:="", PasswordTemplate:="", WritePasswordDocument:="", _
            WritePasswordTemplate:="", Revert:=False, Format:=0, _
            Connection:="yourNamedRange", SQLStatement:="SELECT * FROM `yourNamedRange`", SQLStatement1:=""
        .Destination = 0
        .SuppressBlankLines = True
        .Execute Pause:=False
    End With
    
Application.StatusBar = "Now creating document"

    .ActiveDocument.Content.Font.Name = "Times New Roman"
    '.ActiveDocument.Content.Font.Size = 12

    .ActiveDocument.SaveAs sPath & "\" & findmemo
    .ActiveDocument.Close
    
    .Documents(tempfindmemoname).Close SaveChanges:=False

End With
appWD.Quit
Set appWD = Nothing

'delete temp files
Kill tempfindmemoname
Kill tempsourcename

Application.StatusBar = False
Application.ScreenUpdating = True
Application.Visible = True

MsgBox "A Case Summary Letter has been saved as " & sPath & "\" & findmemo & _
".  Thank you!"

End Sub

Sub NewNeg()

'Create Negative Info Memo for new codes for FFY 2012
Dim appWD As Object
Dim review_number As Long, TempStr As String
Dim sample_month As String, PathStr As String, findmemo As String
Dim countynum As String, districtnum As String, county_name As String

'Find path to Examiner's files on this PC
Set WshNetwork = CreateObject("WScript.Network")
Set oDrives = WshNetwork.EnumNetworkDrives

Dim DLetter As String, DUNC As String
Dim thisws As Worksheet, thiswb As Workbook
Dim sPath As String, tempsourcename As String

sPath = ActiveWorkbook.Path

Set thisws = ActiveSheet
Set thiswb = ActiveWorkbook

'determing what kind of review it is by the first number in the review
review_type = Left(thisws.Name, 1)

'Find drive letter for stat network drive
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

PathStr = DLetter
Application.ScreenUpdating = False

Application.StatusBar = "Starting mail merge ..."

'Copy template data source to active directory
tempsourcename = sPath & "\FM DS Temp.xlsx"
SrceFile = PathStr & "Finding Memo\Finding Memo Data Source.xlsx"
DestFile = tempsourcename
FileCopy SrceFile, DestFile

'Open spreadsheet where all the data is stored to populate findings memo
Workbooks.Open fileName:=tempsourcename, UpdateLinks:=False

Dim datasourcewb As Workbook, datasourcews As Worksheet

Set datasourcewb = ActiveWorkbook
Set datasourcews = ActiveSheet

'deleting information in Finding Memo Data Source spreadsheet
Range("B2:E2") = ""
Range("G2") = ""
Range("O2:S2") = ""
Range("W2:Y2") = ""

'Client Information
Range("B2") = StrConv(thisws.Range("Client_Name"), vbProperCase)
Range("C2") = Right(thisws.Range("Local_Agency"), 2) & "/" & _
    Left(thisws.Range("Case_Record_Number"), 9)
    
    

' Get county number and name in correct format
countynum = Right(thisws.Range("Local_Agency"), 2)
If review_type = "2" Or review_type = "1" Or review_type = "9" Then
districtnum = thisws.Range("District_Code")
Else
districtnum = thisws.Range("District_Code1") & thisws.Range("District_Code2")
End If
'If MA Neg or SNAP neg (because no district on schedule) then lookup county name in temp file
If review_type = "8" Or review_type = "6" Then
   If countynum = "02" Or countynum = "23" Or countynum = "40" Or _
        countynum = "51" Or countynum = "63" Or countynum = "65" Then
        For jj = 50 To 139
            If Cells(jj, 1) = Val(countynum) And Cells(jj, 2) = districtnum Then
                county_name = Cells(jj, 4)
                Exit For
            End If
        Next jj
    Else
        For jj = 50 To 139
            If Cells(jj, 1) = Val(countynum) Then
                county_name = Cells(jj, 4) & " County"
                Exit For
            End If
        Next jj
    End If
Else
    county_name = thisws.Range("County")
End If

    countylookup = countyformat(countynum, county_name, districtnum)
    Range("AK2") = thisws.Range("M5") & " Assistance Office"
    'Look for counties with districts
    'If countynum = "02" Or countynum = "23" Or countynum = "40" Or _
    '    countynum = "51" Or countynum = "63" Or countynum = "65" Then
    '    Range("AK2") = county_name & " Office"
    'Else
    '    Range("AK2") = county_name
    'End If

'Put into temp data file other data
Range("D2") = thisws.Name
review_number = Range("D2")


If review_type = "6" Then
    Range("A7") = Val(thisws.Range("Examiner_Number"))
Else
    Range("A7") = Val(thisws.Range("Examiner_Number1") & thisws.Range("Examiner_Number2"))
End If
    Range("E7") = review_type
    Range("E2") = thisws.Range("Sample_Month") & "/" & thisws.Range("Sample_Year")
    sample_month = thisws.Range("Sample_Month") & thisws.Range("Sample_Year")


'reformat street address
If review_type <> "1" And review_type <> "6" Then
    If thisws.Range("Client_2ndAddress") <> "" Then
        Range("Z2") = StrConv(thisws.Range("Client_1stAddress"), vbProperCase) & ", " & StrConv(thisws.Range("Client_2ndAddress"), vbProperCase)
    Else
        Range("Z2") = StrConv(thisws.Range("Client_1stAddress"), vbProperCase)
    End If
Else
     Range("Z2") = StrConv(thisws.Range("Client_1stAddress"), vbProperCase)
End If
'Reformat city and state
    temparry = Split(thisws.Range("Client_TownStateZip"), ",")
    TempStr = StrConv(Trim(temparry(LBound(temparry))), vbProperCase)

Range("AA2") = TempStr & ", " & Trim(temparry(UBound(temparry)))

'********* to determine if the case would cause an error in 2012... no longer used*******

'If thisws.Range("I49") = 2 Then
'    MsgBox ("New Negative Info Memo not needed for this type of review.  The code on the schedule is a 2.")
'    datasourcewb.Close False
'    Kill tempsourcename
'    Exit Sub
'End If

'If thisws.Range("I49") = "" Then
'    MsgBox ("There is no code on the schedule for this memo.  Please fill either a 1 or 2 in cell.")
'    datasourcewb.Close False
'    Kill tempsourcename
'    Exit Sub
'End If


'Going out to website with list of Executive Directors
        FullTextFileName = datasourcews.Range("U6")
            
            Workbooks.Open fileName:=FullTextFileName, UpdateLinks:=False, ReadOnly:=True
            
            Range("A1:D1000").UnMerge

            'Allegheny Districts are formatted as dates, change them to 02/XX format
            For irow = 1 To 100
                If IsDate(Range("A" & irow)) Then
                    TempStr = WorksheetFunction.Text(Range("A" & irow), "mm/d/yyyy")
                    Range("A" & irow).NumberFormat = "@"
                    Range("A" & irow) = Left(TempStr, 4)
                End If
            Next irow
            
            'Look for county or district number
                found_row = Range("A1:A1000").Find(countylookup).Row
            
            datasourcews.Range("Y2") = Range("D" & found_row)
            'For irow = 1 To 1000
            '    If Range("A" & irow) = countylookup Then
                    'Get ED name
            '        datasourcews.Range("Y2") = Range("D" & irow)
            '        'Get CAO/DO Address
                    caoaddress = ""
                    For ii = 0 To 4
                        If Range("C" & found_row + ii) = "" Then
                            Exit For
                        End If
                        caoaddress = caoaddress & Range("C" & found_row + ii) & ", "
                    Next ii
                    datasourcews.Range("AL2") = caoaddress
                'End If
            'Next irow
            ActiveWorkbook.Close False
            
    
    datasourcewb.Close True
    Application.DisplayAlerts = True
            
    findmemo = "Negative Review Information Memo for Review Number " & review_number & " for Sample Month " & sample_month & ".doc"

    'Copy template cao appointment memo to active directory
    SrceFile = PathStr & "Finding Memo\Negative Review Information Memo Master.docx"
    tempfindmemoname = sPath & "\FM temp.doc"
    DestFile = tempfindmemoname
    FileCopy SrceFile, DestFile

    Set appWD = CreateObject("Word.Application")
    appWD.Visible = True

    With appWD
    .Documents.Open fileName:=tempfindmemoname
    
    'Merge CAO information into template
    With .ActiveDocument.MailMerge
        .OpenDataSource Name:=(tempsourcename), _
            ReadOnly:=True, LinkToSource:=0, AddToRecentFiles:=False, _
            PasswordDocument:="", PasswordTemplate:="", WritePasswordDocument:="", _
            WritePasswordTemplate:="", Revert:=False, Format:=0, _
            Connection:="yourNamedRange", SQLStatement:="SELECT * FROM `yourNamedRange`", SQLStatement1:=""
        .Destination = 0
        .SuppressBlankLines = True
        .Execute Pause:=False
    End With
    
Application.StatusBar = "Now creating document"

    .ActiveDocument.Content.Font.Name = "Times New Roman"
    '.ActiveDocument.Content.Font.Size = 12

    .ActiveDocument.SaveAs sPath & "\" & findmemo
    .ActiveDocument.Close
    
    .Documents(tempfindmemoname).Close SaveChanges:=False

End With
appWD.Quit
Set appWD = Nothing

'delete temp files
Kill tempfindmemoname
Kill tempsourcename

Application.StatusBar = False
Application.ScreenUpdating = True
Application.Visible = True

MsgBox "A Negative Review Information Memo has been saved as " & sPath & "\" & findmemo & _
"."

End Sub
Sub ComSpouse()

'Create CAO Appointment Memo
Dim appWD As Object
Dim review_number As Long, TempStr As String
Dim sample_month As String, PathStr As String, findmemo As String
Dim countynum As String, districtnum As String

'Find path to Examiner's files on this PC
Set WshNetwork = CreateObject("WScript.Network")
Set oDrives = WshNetwork.EnumNetworkDrives

Dim DLetter As String, DUNC As String
Dim thisws As Worksheet, thiswb As Workbook
Dim sPath As String, tempsourcename As String

sPath = ActiveWorkbook.Path

Set thisws = ActiveSheet
Set thiswb = ActiveWorkbook

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

PathStr = DLetter
Application.ScreenUpdating = False

'Application.StatusBar = "Starting mail merge ..."

'Copy template data source to active directory
tempsourcename = sPath & "\FM DS Temp.xlsx"
SrceFile = PathStr & "Finding Memo\Finding Memo Data Source.xlsx"
DestFile = tempsourcename
FileCopy SrceFile, DestFile

'Copy community spouse memo
    SrceFile = PathStr & "Finding Memo\Community Spouse Questionaire.xlsm"
    tempfindmemoname = sPath & "\Community Spouse Questionaire.xlsm"
    DestFile = tempfindmemoname
    FileCopy SrceFile, DestFile

'Open spreadsheet where all the data is stored to populate findings memo
Workbooks.Open fileName:=tempfindmemoname, UpdateLinks:=False

Dim datasourcewb As Workbook, datasourcews As Worksheet

Set datasourcewb = ActiveWorkbook
Set datasourcews = ActiveSheet

' Get county number and name in correct format
countynum = Right(thisws.Range("AC10"), 2)
districtnum = thisws.Range("AH10")
countylookup = countyformat(countynum, thisws.Range("O4"), districtnum)

'Look for counties with districts
If countynum = "02" Or countynum = "23" Or countynum = "40" Or _
        countynum = "51" Or countynum = "63" Or countynum = "65" Then
    Range("B46") = thisws.Range("O4")
Else
    Range("B46") = thisws.Range("O4")
End If

'Client Information
Range("K13") = StrConv(thisws.Range("B2"), vbProperCase)
Range("C44") = StrConv(thisws.Range("B2"), vbProperCase)
'County and Case Number
Range("K14") = thisws.Range("H10")
Range("G46") = thisws.Range("H10")

'review number
Range("K46") = thisws.Range("A10") & thisws.Range("B10") & thisws.Range("c10") & thisws.Range("D10") & thisws.Range("E10") & thisws.Range("F10")
review_number = Range("J46")
'review month
Range("H44") = thisws.Range("AM10")
sample_month = Range("H44")
'supervisor/examiner
Range("G36") = thisws.Range("AF2")
Range("E7") = review_type



Workbooks.Open fileName:=tempsourcename, UpdateLinks:=False

'reviewer number
Range("A7") = Val(thisws.Range("AO3") & thisws.Range("AP3"))
datasourcews.Range("G36") = Range("AM2")
datasourcews.Range("K9") = Range("AN2")
datasourcews.Range("K10") = Range("AF2")
datasourcews.Range("K11") = Range("AG2")
datasourcews.Range("A4") = "Division of Quality Control, " & Range("AC2") & "Field Office"
datasourcews.Range("A5") = Range("AD2")
datasourcews.Range("A6") = Range("AE2")

'due date
datasourcews.Range("A29") = Range("DI2")

ActiveWorkbook.Close False

sample_month = Application.WorksheetFunction.Text(sample_month, "YYYYMM")
findmemo = "\Community Spouse Questionaire for Review Number " & _
    review_number & " for Sample Month " & sample_month & ".xlsm"
Workbooks("Community Spouse Questionaire.xlsm").Activate
ActiveWorkbook.SaveAs fileName:=sPath & findmemo, _
           FileFormat:=52, _
           Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, _
           CreateBackup:=False
ActiveWorkbook.Close False

'delete temp files
Kill tempsourcename
Kill tempfindmemoname

Application.StatusBar = False
Application.ScreenUpdating = True
Application.Visible = True

MsgBox "A Community Spouse form has been saved as " & sPath & "\" & tempfindmemoname & _
". - Thank you!"

End Sub

Sub QC14C()

'Create QC 14C Form
Dim appWD As Object
Dim review_number As Long
Dim sample_month As String, PathStr As String, findmemo As String
Dim countynum As String, districtnum As String

'Find path to Examiner's files on this PC
Set WshNetwork = CreateObject("WScript.Network")
Set oDrives = WshNetwork.EnumNetworkDrives

Dim DLetter As String, DUNC As String
Dim thisws As Worksheet, thiswb As Workbook
Dim sPath As String, tempsourcename As String

sPath = ActiveWorkbook.Path

Set thisws = ActiveSheet
Set thiswb = ActiveWorkbook

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

PathStr = DLetter
Application.ScreenUpdating = False

Application.StatusBar = "Starting mail merge ..."

'Copy template data source to active directory
tempsourcename = sPath & "\FM DS Temp.xlsx"
SrceFile = PathStr & "Finding Memo\Finding Memo Data Source.xlsx"
DestFile = tempsourcename
FileCopy SrceFile, DestFile

'Open spreadsheet where all the data is stored to populate findings memo
Workbooks.Open fileName:=tempsourcename, UpdateLinks:=False

Dim datasourcewb As Workbook, datasourcews As Worksheet

Set datasourcewb = ActiveWorkbook
Set datasourcews = ActiveSheet


'deleting information in Finding Memo Data Source spreadsheet
Range("B2:E2") = ""
Range("G2") = ""
Range("O2:S2") = ""
Range("W2:Y2") = ""

'determing what kind of review it is by the first number in the review
review_type = Left(thisws.Name, 1)

'SNAP Positive
If review_type = 5 Then

'Client Information
Range("B2") = StrConv(thisws.Range("B4"), vbProperCase)
Range("C2") = thisws.Range("C155") & thisws.Range("D155") & "/" & Left(thisws.Range("I18"), 9)

' Get county number and name in correct format
countynum = thisws.Range("C155") & thisws.Range("D155")
districtnum = thisws.Range("E155") & thisws.Range("F155")
countylookup = countyformat(countynum, thisws.Range("M5"), districtnum)
Range("AK2") = thisws.Range("M5") & " Assistance Office"
'Look for counties with districts
'If countynum = "02" Or countynum = "23" Or countynum = "40" Or _
'        countynum = "51" Or countynum = "63" Or countynum = "65" Then
'    Range("AK2") = thisws.Range("M5") & " Office"
'Else
'    Range("AK2") = thisws.Range("M5") & " CAO"
'End If

'Put into temp data file other data
Range("D2") = thisws.Range("A18")
review_number = Range("D2")
Range("E2") = thisws.Range("AD18") & "/" & thisws.Range("AG18")
sample_month = thisws.Range("AD18") & thisws.Range("AG18")
Range("A7") = Val(thisws.Range("AJ5") & thisws.Range("AK5"))
Range("G2") = Range("C7") & "/" & thisws.Range("AE5")
Range("E7") = review_type
Range("O2") = thiswb.Sheets("FS Computation").Range("B60")
Range("P2") = thiswb.Sheets("FS Computation").Range("C60")
Range("Q2") = thisws.Range("Y22")
Range("R2") = thisws.Range("B29") & " - " & thisws.Range("G29") & " - " & thisws.Range("K29")
Range("E17") = Val(thisws.Range("K22"))
'reformat street address
If thisws.Range("B6") <> "" Then
    Range("Z2") = StrConv(thisws.Range("B5"), vbProperCase) & ", " & StrConv(thisws.Range("B6"), vbProperCase)
Else
    Range("Z2") = StrConv(thisws.Range("B5"), vbProperCase)
End If

'Reformat city and state
    temparry = Split(thisws.Range("B7"), ",")
    TempStr = StrConv(Trim(temparry(LBound(temparry))), vbProperCase)

Range("AA2") = TempStr & ", " & Trim(temparry(UBound(temparry)))


ElseIf review_type = 2 Then

'Client Information
Range("B2") = StrConv(thisws.Range("B2"), vbProperCase) 'client name
Range("C2") = Right(thisws.Range("AC10"), 2) & "/" & Left(thisws.Range("H10"), 9) 'county and case number

' Get county number and name in correct format
countynum = Right(thisws.Range("AC10"), 2) ' county
districtnum = thisws.Range("AH10") 'district
countylookup = countyformat(countynum, thisws.Range("O4"), districtnum) ' County Name
Range("AK2") = thisws.Range("O4") & " Assistance Office"
'Look for counties with districts
'If countynum = "02" Or countynum = "23" Or countynum = "40" Or _
'        countynum = "51" Or countynum = "63" Or countynum = "65" Then
'    Range("AK2") = thisws.Range("M5") & " Office"
'Else
'    Range("AK2") = thisws.Range("M5") & " CAO"
'End If

'Put into temp data file other data
Range("D2") = thisws.Range("a10") & thisws.Range("b10") & thisws.Range("c10") & thisws.Range("d10") & thisws.Range("e10") & thisws.Range("f10") 'Review Number
review_number = Range("D2")
Range("E2") = thisws.Range("AM10") 'Sample Month
sample_month = thisws.Range("AM10") 'Sample Month
Range("A7") = Val(thisws.Range("AO3") & thisws.Range("AP3")) 'Examiner Number
Range("G2") = Range("C7") & "/" & thisws.Range("AF2") 'Examiner Name
Range("E7") = review_type
'Range("E17") = Val(thisws.Range("K22"))
'reformat street address
If thisws.Range("B4") <> "" Then
    Range("Z2") = StrConv(thisws.Range("B3"), vbProperCase) & ", " & StrConv(thisws.Range("B4"), vbProperCase)
Else
    Range("Z2") = StrConv(thisws.Range("B3"), vbProperCase)
End If

'Reformat city and state
    temparry = Split(thisws.Range("B5"), ",")
    TempStr = StrConv(Trim(temparry(LBound(temparry))), vbProperCase)

Range("AA2") = TempStr & ", " & Trim(temparry(UBound(temparry)))

Else

    MsgBox ("QC14C Request Letter not needed for this type of review")

End

End If

'Going out to website with list of Executive Directors
        FullTextFileName = datasourcews.Range("U6")
            
            'ActiveWorkbook.FollowHyperlink (FullTextFileName)
            Workbooks.Open fileName:=FullTextFileName, UpdateLinks:=False, ReadOnly:=True
            
            Range("A1:D1000").UnMerge

            'Allegheny Districts are formatted as dates, change them to 02/XX format
            For irow = 1 To 100
                If IsDate(Range("A" & irow)) Then
                    TempStr = WorksheetFunction.Text(Range("A" & irow), "mm/d/yyyy")
                    Range("A" & irow).NumberFormat = "@"
                    Range("A" & irow) = Left(TempStr, 4)
                End If
            Next irow
            
            'Look for county or district number
                found_row = Range("A1:A1000").Find(countylookup).Row
            
            datasourcews.Range("Y2") = Range("D" & found_row)
            'For irow = 1 To 1000
            '    If Range("A" & irow) = countylookup Then
                    'Get ED name
            '        datasourcews.Range("Y2") = Range("D" & irow)
            '        'Get CAO/DO Address
                    caoaddress = ""
                    For ii = 0 To 4
                        If Range("C" & found_row + ii) = "" Then
                            Exit For
                        End If
                        caoaddress = caoaddress & Range("C" & found_row + ii) & ", "
                    Next ii
                    datasourcews.Range("AL2") = caoaddress
                'End If
            'Next irow
            ActiveWorkbook.Close False
            
            'Application.Wait Now + TimeSerial(0, 0, 2)
            'CloseApp "CAO List", ""

            datasourcewb.Close True
            
             If review_type = 2 Then
                sample_month = Left(sample_month, 2) & Right(sample_month, 4)
                findmemo = "QC FINDING Review Number " & review_number & " for Sample Month " & sample_month & ".doc"
             Else
                findmemo = "QC FINDING Review Number " & review_number & " for Sample Month " & sample_month & ".doc"
             End If

    If review_type = 2 Then
        findmemo = "QC14C Coop Memo for Review Number " & review_number & " for Sample Month " & sample_month & ".doc"
    Else
        findmemo = "QC14C Request Letter for Review Number " & review_number & " for Sample Month " & sample_month & ".doc"
    End If
    'Copy template finding memo to active directory
    SrceFile = PathStr & "Finding Memo\SNAP QC14C.docx"
    tempfindmemoname = sPath & "\FM temp.doc"
    DestFile = tempfindmemoname
    FileCopy SrceFile, DestFile

    Set appWD = CreateObject("Word.Application")
    appWD.Visible = True

    With appWD
    .Documents.Open fileName:=tempfindmemoname
    
    With .ActiveDocument.MailMerge
        .OpenDataSource Name:=(tempsourcename), _
            ReadOnly:=True, LinkToSource:=0, AddToRecentFiles:=False, _
            PasswordDocument:="", PasswordTemplate:="", WritePasswordDocument:="", _
            WritePasswordTemplate:="", Revert:=False, Format:=0, _
            Connection:="yourNamedRange", SQLStatement:="SELECT * FROM `yourNamedRange`", SQLStatement1:=""
        .Destination = 0
        .SuppressBlankLines = True
        .Execute Pause:=False
    End With
    
    '.ActiveDocument.Protect Type:=2, NoReset:=True
    
Application.StatusBar = "Now creating document"

    .ActiveDocument.Content.Font.Name = "Times New Roman"
    '.ActiveDocument.Content.Font.Size = 12

    .ActiveDocument.SaveAs sPath & "\" & findmemo
    .ActiveDocument.Close
    
    .Documents(tempfindmemoname).Close SaveChanges:=False

End With
appWD.Quit
Set appWD = Nothing

'delete temp files
Kill tempfindmemoname
Kill tempsourcename

Application.StatusBar = False
Application.ScreenUpdating = True
Application.Visible = True

MsgBox "A QC14C Request Letter has been saved as " & sPath & "\" & findmemo & _
". Please close the Internet Explorer with the CAO Directory. - Thank you!"

End Sub
Sub QC14()

'Create QC 14C Form
Dim appWD As Object
Dim review_number As Long
Dim sample_month As String, PathStr As String, findmemo As String
Dim countynum As String, districtnum As String

'Find path to Examiner's files on this PC
Set WshNetwork = CreateObject("WScript.Network")
Set oDrives = WshNetwork.EnumNetworkDrives

Dim DLetter As String, DUNC As String
Dim thisws As Worksheet, thiswb As Workbook
Dim sPath As String, tempsourcename As String

sPath = ActiveWorkbook.Path

Set thisws = ActiveSheet
Set thiswb = ActiveWorkbook

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

PathStr = DLetter
Application.ScreenUpdating = False

Application.StatusBar = "Starting mail merge ..."

'Copy template data source to active directory
tempsourcename = sPath & "\FM DS Temp.xlsx"
SrceFile = PathStr & "Finding Memo\Finding Memo Data Source.xlsx"
DestFile = tempsourcename
FileCopy SrceFile, DestFile

'Open spreadsheet where all the data is stored to populate findings memo
Workbooks.Open fileName:=tempsourcename, UpdateLinks:=False

Dim datasourcewb As Workbook, datasourcews As Worksheet

Set datasourcewb = ActiveWorkbook
Set datasourcews = ActiveSheet


'deleting information in Finding Memo Data Source spreadsheet
Range("B2:E2") = ""
Range("G2") = ""
Range("O2:S2") = ""
Range("W2:Y2") = ""

'determing what kind of review it is by the first number in the review
review_type = Left(thisws.Name, 1)

'SNAP Positive
If review_type = 5 Then

'Client Information
Range("B2") = StrConv(thisws.Range("B4"), vbProperCase)
Range("C2") = thisws.Range("C155") & thisws.Range("D155") & "/" & Left(thisws.Range("I18"), 9)

' Get county number and name in correct format
countynum = thisws.Range("C155") & thisws.Range("D155")
districtnum = thisws.Range("E155") & thisws.Range("F155")
countylookup = countyformat(countynum, thisws.Range("M5"), districtnum)
Range("AK2") = thisws.Range("M5") & " Assistance Office"
'Look for counties with districts
'If countynum = "02" Or countynum = "23" Or countynum = "40" Or _
'        countynum = "51" Or countynum = "63" Or countynum = "65" Then
'    Range("AK2") = thisws.Range("M5") & " Office"
'Else
'    Range("AK2") = thisws.Range("M5") & " CAO"
'End If

'Put into temp data file other data
Range("D2") = thisws.Range("A18")
review_number = Range("D2")
Range("E2") = thisws.Range("AD18") & "/" & thisws.Range("AG18")
sample_month = thisws.Range("AD18") & thisws.Range("AG18")
Range("A7") = Val(thisws.Range("AJ5") & thisws.Range("AK5"))
Range("G2") = Range("C7") & "/" & thisws.Range("AE5")
Range("E7") = review_type
Range("O2") = thiswb.Sheets("FS Computation").Range("B60")
Range("P2") = thiswb.Sheets("FS Computation").Range("C60")
Range("Q2") = thisws.Range("Y22")
Range("R2") = thisws.Range("B29") & " - " & thisws.Range("G29") & " - " & thisws.Range("K29")
Range("E17") = Val(thisws.Range("K22"))
'reformat street address
If thisws.Range("B6") <> "" Then
    Range("Z2") = StrConv(thisws.Range("B5"), vbProperCase) & ", " & StrConv(thisws.Range("B6"), vbProperCase)
Else
    Range("Z2") = StrConv(thisws.Range("B5"), vbProperCase)
End If

'Reformat city and state
    temparry = Split(thisws.Range("B7"), ",")
    TempStr = StrConv(Trim(temparry(LBound(temparry))), vbProperCase)

Range("AA2") = TempStr & ", " & Trim(temparry(UBound(temparry)))


ElseIf review_type = 2 Then

'Client Information
Range("B2") = StrConv(thisws.Range("B2"), vbProperCase) 'client name
Range("C2") = Right(thisws.Range("AC10"), 2) & "/" & Left(thisws.Range("H10"), 9) 'county and case number

' Get county number and name in correct format
countynum = Right(thisws.Range("AC10"), 2) ' county
districtnum = thisws.Range("AH10") 'district
countylookup = countyformat(countynum, thisws.Range("O4"), districtnum) ' County Name
Range("AK2") = thisws.Range("O4") & " Assistance Office"
'Look for counties with districts
'If countynum = "02" Or countynum = "23" Or countynum = "40" Or _
'        countynum = "51" Or countynum = "63" Or countynum = "65" Then
'    Range("AK2") = thisws.Range("M5") & " Office"
'Else
'    Range("AK2") = thisws.Range("M5") & " CAO"
'End If

'Put into temp data file other data
Range("D2") = thisws.Range("a10") & thisws.Range("b10") & thisws.Range("c10") & thisws.Range("d10") & thisws.Range("e10") & thisws.Range("f10") 'Review Number
review_number = Range("D2")
Range("E2") = thisws.Range("AM10") 'Sample Month
sample_month = thisws.Range("AM10") 'Sample Month
Range("A7") = Val(thisws.Range("AO3") & thisws.Range("AP3")) 'Examiner Number
Range("G2") = Range("C7") & "/" & thisws.Range("AF2") 'Examiner Name
Range("E7") = review_type
'Range("E17") = Val(thisws.Range("K22"))
'reformat street address
If thisws.Range("B4") <> "" Then
    Range("Z2") = StrConv(thisws.Range("B3"), vbProperCase) & ", " & StrConv(thisws.Range("B4"), vbProperCase)
Else
    Range("Z2") = StrConv(thisws.Range("B3"), vbProperCase)
End If

'Reformat city and state
    temparry = Split(thisws.Range("B5"), ",")
    TempStr = StrConv(Trim(temparry(LBound(temparry))), vbProperCase)

Range("AA2") = TempStr & ", " & Trim(temparry(UBound(temparry)))

Else

    MsgBox ("QC14 Request Letter not needed for this type of review")

End

End If

'Going out to website with list of Executive Directors
        FullTextFileName = datasourcews.Range("U6")
            
            'ActiveWorkbook.FollowHyperlink (FullTextFileName)
            Workbooks.Open fileName:=FullTextFileName, UpdateLinks:=False, ReadOnly:=True
            
            Range("A1:D1000").UnMerge

            'Allegheny Districts are formatted as dates, change them to 02/XX format
            For irow = 1 To 100
                If IsDate(Range("A" & irow)) Then
                    TempStr = WorksheetFunction.Text(Range("A" & irow), "mm/d/yyyy")
                    Range("A" & irow).NumberFormat = "@"
                    Range("A" & irow) = Left(TempStr, 4)
                End If
            Next irow
            
            'Look for county or district number
                found_row = Range("A1:A1000").Find(countylookup).Row
            
            datasourcews.Range("Y2") = Range("D" & found_row)
            'For irow = 1 To 1000
            '    If Range("A" & irow) = countylookup Then
                    'Get ED name
            '        datasourcews.Range("Y2") = Range("D" & irow)
            '        'Get CAO/DO Address
                    caoaddress = ""
                    For ii = 0 To 4
                        If Range("C" & found_row + ii) = "" Then
                            Exit For
                        End If
                        caoaddress = caoaddress & Range("C" & found_row + ii) & ", "
                    Next ii
                    datasourcews.Range("AL2") = caoaddress
                'End If
            'Next irow
            ActiveWorkbook.Close False
            
            'Application.Wait Now + TimeSerial(0, 0, 2)
            'CloseApp "CAO List", ""
            
        'put in deadline - 2 weeks
        datasourcews.Range("AQ2") = datasourcews.Range("A2") + 7
        datasourcewb.Close True
            
             If review_type = 2 Then
                sample_month = Left(sample_month, 2) & Right(sample_month, 4)
                findmemo = "QC14 CAO Request Memo for Review Number " & review_number & " for Sample Month " & sample_month & ".doc"
             Else
                findmemo = "QC14 CAO Request Memo for Review Number " & review_number & " for Sample Month " & sample_month & ".doc"
             End If

    'If review_type = 2 Then
    '    findmemo = "QC14 CAO Request Memo for Review Number " & review_number & " for Sample Month " & sample_month & ".docx"
    'Else
    '    findmemo = "QC14 Request Letter for Review Number " & review_number & " for Sample Month " & sample_month & ".doc"
    'End If
    'Copy template finding memo to active directory
    SrceFile = PathStr & "Finding Memo\MA QC 14 Request for CAO Assistance.docx"
    tempfindmemoname = sPath & "\FM temp.docx"
    DestFile = tempfindmemoname
    FileCopy SrceFile, DestFile

    Set appWD = CreateObject("Word.Application")
    appWD.Visible = True

    With appWD
    .Documents.Open fileName:=tempfindmemoname
    
    With .ActiveDocument.MailMerge
        .OpenDataSource Name:=(tempsourcename), _
            ReadOnly:=True, LinkToSource:=0, AddToRecentFiles:=False, _
            PasswordDocument:="", PasswordTemplate:="", WritePasswordDocument:="", _
            WritePasswordTemplate:="", Revert:=False, Format:=0, _
            Connection:="yourNamedRange", SQLStatement:="SELECT * FROM `yourNamedRange`", SQLStatement1:=""
        .Destination = 0
        .SuppressBlankLines = True
        .Execute Pause:=False
    End With
    
    '.ActiveDocument.Protect Type:=2, NoReset:=True
    
Application.StatusBar = "Now creating document"

    .ActiveDocument.Content.Font.Name = "Times New Roman"
    '.ActiveDocument.Content.Font.Size = 12

    .ActiveDocument.SaveAs sPath & "\" & findmemo
    .ActiveDocument.Close
    
    .Documents(tempfindmemoname).Close SaveChanges:=False

End With
appWD.Quit
Set appWD = Nothing

'delete temp files
Kill tempfindmemoname
Kill tempsourcename

Application.StatusBar = False
Application.ScreenUpdating = True
Application.Visible = True

MsgBox "A QC14 Request Letter has been saved as " & sPath & "\" & findmemo & _
". If needed, please close the Internet Explorer with the CAO Directory. - Thank you!"

End Sub

Sub Prelim_Info()

'Create QC Preliminary Information Memo
Dim appWD As Object
Dim review_number As Long
Dim sample_month As String, PathStr As String, findmemo As String
Dim countynum As String, districtnum As String

'Find path to Examiner's files on this PC
Set WshNetwork = CreateObject("WScript.Network")
Set oDrives = WshNetwork.EnumNetworkDrives

Dim DLetter As String, DUNC As String
Dim thisws As Worksheet, thiswb As Workbook
Dim sPath As String, tempsourcename As String

sPath = ActiveWorkbook.Path

Set thisws = ActiveSheet
Set thiswb = ActiveWorkbook

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

PathStr = DLetter
Application.ScreenUpdating = False

Application.StatusBar = "Starting mail merge ..."

'Copy template data source to active directory
tempsourcename = sPath & "\FM DS Temp.xlsx"
SrceFile = PathStr & "Finding Memo\Finding Memo Data Source.xlsx"
DestFile = tempsourcename
FileCopy SrceFile, DestFile

'Open spreadsheet where all the data is stored to populate findings memo
Workbooks.Open fileName:=tempsourcename, UpdateLinks:=False

Dim datasourcewb As Workbook, datasourcews As Worksheet

Set datasourcewb = ActiveWorkbook
Set datasourcews = ActiveSheet


'deleting information in Finding Memo Data Source spreadsheet
Range("B2:E2") = ""
Range("G2") = ""
Range("O2:S2") = ""
Range("W2:Y2") = ""

'determing what kind of review it is by the first number in the review
review_type = Left(thisws.Name, 1)

'SNAP Positive
If review_type = 5 Then

'Client Information
Range("B2") = StrConv(thisws.Range("B4"), vbProperCase)
Range("C2") = thisws.Range("C155") & thisws.Range("D155") & "/" & Left(thisws.Range("I18"), 9)

' Get county number and name in correct format
countynum = thisws.Range("C155") & thisws.Range("D155")
districtnum = thisws.Range("E155") & thisws.Range("F155")
countylookup = countyformat(countynum, thisws.Range("M5"), districtnum)
Range("AK2") = thisws.Range("M5") & " Assistance Office"
'Look for counties with districts
'If countynum = "02" Or countynum = "23" Or countynum = "40" Or _
'        countynum = "51" Or countynum = "63" Or countynum = "65" Then
'    Range("AK2") = thisws.Range("M5") & " Office"
'Else
'    Range("AK2") = thisws.Range("M5") & " CAO"
'End If

'Put into temp data file other data
Range("D2") = thisws.Range("A18")
review_number = Range("D2")
Range("E2") = thisws.Range("AD18") & "/" & thisws.Range("AG18")
sample_month = thisws.Range("AD18") & thisws.Range("AG18")
Range("A7") = Val(thisws.Range("AJ5") & thisws.Range("AK5"))
Range("G2") = Range("C7") & "/" & thisws.Range("AE5")
Range("E7") = review_type
Range("O2") = thiswb.Sheets("FS Computation").Range("B60")
Range("P2") = thiswb.Sheets("FS Computation").Range("C60")
Range("Q2") = thisws.Range("Y22")
Range("R2") = thisws.Range("B29") & " - " & thisws.Range("G29") & " - " & thisws.Range("K29")
Range("E17") = Val(thisws.Range("K22"))
'reformat street address
If thisws.Range("B6") <> "" Then
    Range("Z2") = StrConv(thisws.Range("B5"), vbProperCase) & ", " & StrConv(thisws.Range("B6"), vbProperCase)
Else
    Range("Z2") = StrConv(thisws.Range("B5"), vbProperCase)
End If

'Reformat city and state
    temparry = Split(thisws.Range("B7"), ",")
    TempStr = StrConv(Trim(temparry(LBound(temparry))), vbProperCase)

Range("AA2") = TempStr & ", " & Trim(temparry(UBound(temparry)))


ElseIf review_type = 2 Then

'Client Information
Range("B2") = StrConv(thisws.Range("B2"), vbProperCase) 'client name
Range("C2") = Right(thisws.Range("AC10"), 2) & "/" & Left(thisws.Range("H10"), 9) 'county and case number

' Get county number and name in correct format
countynum = Right(thisws.Range("AC10"), 2) ' county
districtnum = thisws.Range("AH10") 'district
countylookup = countyformat(countynum, thisws.Range("O4"), districtnum) ' County Name
Range("AK2") = thisws.Range("O4") & " Assistance Office"
'Look for counties with districts
'If countynum = "02" Or countynum = "23" Or countynum = "40" Or _
'        countynum = "51" Or countynum = "63" Or countynum = "65" Then
'    Range("AK2") = thisws.Range("M5") & " Office"
'Else
'    Range("AK2") = thisws.Range("M5") & " CAO"
'End If

'Put into temp data file other data
Range("D2") = thisws.Range("a10") & thisws.Range("b10") & thisws.Range("c10") & thisws.Range("d10") & thisws.Range("e10") & thisws.Range("f10") 'Review Number
review_number = Range("D2")
Range("E2") = thisws.Range("AM10") 'Sample Month
sample_month = thisws.Range("AM10") 'Sample Month
Range("A7") = Val(thisws.Range("AO3") & thisws.Range("AP3")) 'Examiner Number
Range("G2") = Range("C7") & "/" & thisws.Range("AF2") 'Examiner Name
Range("E7") = review_type
'Range("E17") = Val(thisws.Range("K22"))
'reformat street address
If thisws.Range("B4") <> "" Then
    Range("Z2") = StrConv(thisws.Range("B3"), vbProperCase) & ", " & StrConv(thisws.Range("B4"), vbProperCase)
Else
    Range("Z2") = StrConv(thisws.Range("B3"), vbProperCase)
End If

'Reformat city and state
    temparry = Split(thisws.Range("B5"), ",")
    TempStr = StrConv(Trim(temparry(LBound(temparry))), vbProperCase)
    Range("AA2") = TempStr & ", " & Trim(temparry(UBound(temparry)))
    
Range("CE2") = thisws.Range("Q10") 'category
Range("CG2") = thisws.Range("Z10") 'grant group

Else

    MsgBox ("Preliminary Information Memo is not needed for this type of review")

End

End If

'Going out to website with list of Executive Directors
        FullTextFileName = datasourcews.Range("U6")
            
            'ActiveWorkbook.FollowHyperlink (FullTextFileName)
            Workbooks.Open fileName:=FullTextFileName, UpdateLinks:=False, ReadOnly:=True
            
            Range("A1:D1000").UnMerge

            'Allegheny Districts are formatted as dates, change them to 02/XX format
            For irow = 1 To 100
                If IsDate(Range("A" & irow)) Then
                    TempStr = WorksheetFunction.Text(Range("A" & irow), "mm/d/yyyy")
                    Range("A" & irow).NumberFormat = "@"
                    Range("A" & irow) = Left(TempStr, 4)
                End If
            Next irow
            
            'Look for county or district number
                found_row = Range("A1:A1000").Find(countylookup).Row
            
            datasourcews.Range("Y2") = Range("D" & found_row)
            'For irow = 1 To 1000
            '    If Range("A" & irow) = countylookup Then
                    'Get ED name
            '        datasourcews.Range("Y2") = Range("D" & irow)
            '        'Get CAO/DO Address
                    caoaddress = ""
                    For ii = 0 To 4
                        If Range("C" & found_row + ii) = "" Then
                            Exit For
                        End If
                        caoaddress = caoaddress & Range("C" & found_row + ii) & ", "
                    Next ii
                    datasourcews.Range("AL2") = caoaddress
                'End If
            'Next irow
            ActiveWorkbook.Close False
            
            'Application.Wait Now + TimeSerial(0, 0, 2)
            'CloseApp "CAO List", ""
            
        'put in deadline - 2 weeks
        datasourcews.Range("AQ2") = datasourcews.Range("A2") + 7
        datasourcewb.Close True
            
             If review_type = 2 Then
                sample_month = Left(sample_month, 2) & Right(sample_month, 4)
                findmemo = "Preliminary Information Memo for Review Number " & review_number & " for Sample Month " & sample_month & ".doc"
             Else
                findmemo = "Preliminary Information  Memo for Review Number " & review_number & " for Sample Month " & sample_month & ".doc"
             End If

    'If review_type = 2 Then
    '    findmemo = "QC14 CAO Request Memo for Review Number " & review_number & " for Sample Month " & sample_month & ".docx"
    'Else
    '    findmemo = "QC14 Request Letter for Review Number " & review_number & " for Sample Month " & sample_month & ".doc"
    'End If
    'Copy template finding memo to active directory
    SrceFile = PathStr & "Finding Memo\MA Preliminary Info Memo.docx"
    tempfindmemoname = sPath & "\FM temp.docx"
    DestFile = tempfindmemoname
    FileCopy SrceFile, DestFile

    Set appWD = CreateObject("Word.Application")
    appWD.Visible = True

    With appWD
    .Documents.Open fileName:=tempfindmemoname
    
    With .ActiveDocument.MailMerge
        .OpenDataSource Name:=(tempsourcename), _
            ReadOnly:=True, LinkToSource:=0, AddToRecentFiles:=False, _
            PasswordDocument:="", PasswordTemplate:="", WritePasswordDocument:="", _
            WritePasswordTemplate:="", Revert:=False, Format:=0, _
            Connection:="yourNamedRange", SQLStatement:="SELECT * FROM `yourNamedRange`", SQLStatement1:=""
        .Destination = 0
        .SuppressBlankLines = True
        .Execute Pause:=False
    End With
    
    '.ActiveDocument.Protect Type:=2, NoReset:=True
    
Application.StatusBar = "Now creating document"

    .ActiveDocument.Content.Font.Name = "Times New Roman"
    '.ActiveDocument.Content.Font.Size = 12

    .ActiveDocument.SaveAs sPath & "\" & findmemo
    .ActiveDocument.Close
    
    .Documents(tempfindmemoname).Close SaveChanges:=False

End With
appWD.Quit
Set appWD = Nothing

'delete temp files
Kill tempfindmemoname
Kill tempsourcename

Application.StatusBar = False
Application.ScreenUpdating = True
Application.Visible = True

MsgBox "A Preliminary Information Letter has been saved as " & sPath & "\" & findmemo & _
". If needed, please close the Internet Explorer with the CAO Directory. - Thank you!"

End Sub
Sub MA_WAIVER()

'Create MA Waiver Memo
Dim appWD As Object
Dim review_number As Long
Dim sample_month As String, PathStr As String, findmemo As String
Dim countynum As String, districtnum As String

'Find path to Examiner's files on this PC
Set WshNetwork = CreateObject("WScript.Network")
Set oDrives = WshNetwork.EnumNetworkDrives

Dim DLetter As String, DUNC As String
Dim thisws As Worksheet, thiswb As Workbook
Dim sPath As String, tempsourcename As String

sPath = ActiveWorkbook.Path

Set thisws = ActiveSheet
Set thiswb = ActiveWorkbook

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

PathStr = DLetter
Application.ScreenUpdating = False

Application.StatusBar = "Starting mail merge ..."

'Copy template data source to active directory
tempsourcename = sPath & "\FM DS Temp.xlsx"
SrceFile = PathStr & "Finding Memo\Finding Memo Data Source.xlsx"
DestFile = tempsourcename
FileCopy SrceFile, DestFile

'Open spreadsheet where all the data is stored to populate findings memo
Workbooks.Open fileName:=tempsourcename, UpdateLinks:=False

Dim datasourcewb As Workbook, datasourcews As Worksheet

Set datasourcewb = ActiveWorkbook
Set datasourcews = ActiveSheet


'deleting information in Finding Memo Data Source spreadsheet
Range("B2:E2") = ""
Range("G2") = ""
Range("O2:S2") = ""
Range("W2:Y2") = ""

'determing what kind of review it is by the first number in the review
review_type = Left(thisws.Name, 1)

'SNAP Positive
If review_type = 5 Then

'Client Information
Range("B2") = StrConv(thisws.Range("B4"), vbProperCase)
Range("C2") = Right(thisws.Range("X18"), 2) & "/" & Left(thisws.Range("I18"), 9)

' Get county number and name in correct format
countynum = Right(thisws.Range("X18"), 2)
districtnum = thisws.Range("B153") & thisws.Range("C153")
countylookup = countyformat(countynum, thisws.Range("M5"), districtnum)
Range("AK2") = thisws.Range("M5") & " Assistance Office"
'Look for counties with districts
'If countynum = "02" Or countynum = "23" Or countynum = "40" Or _
'        countynum = "51" Or countynum = "63" Or countynum = "65" Then
'    Range("AK2") = thisws.Range("M5") & " Office"
'Else
'    Range("AK2") = thisws.Range("M5") & " CAO"
'End If

'Put into temp data file other data
Range("D2") = thisws.Range("A18")
review_number = Range("D2")
Range("E2") = thisws.Range("AD18") & "/" & thisws.Range("AG18")
sample_month = thisws.Range("AD18") & thisws.Range("AG18")
Range("A7") = Val(thisws.Range("AJ5") & thisws.Range("AK5"))
Range("G2") = Range("C7") & "/" & thisws.Range("AE5")
Range("E7") = review_type
Range("O2") = thiswb.Sheets("FS Computation").Range("B60")
Range("P2") = thiswb.Sheets("FS Computation").Range("C60")
Range("Q2") = thisws.Range("Y22")
Range("R2") = thisws.Range("B29") & " - " & thisws.Range("G29") & " - " & thisws.Range("K29")
Range("E17") = Val(thisws.Range("K22"))
'reformat street address
If thisws.Range("B6") <> "" Then
    Range("Z2") = StrConv(thisws.Range("B5"), vbProperCase) & ", " & StrConv(thisws.Range("B6"), vbProperCase)
Else
    Range("Z2") = StrConv(thisws.Range("B5"), vbProperCase)
End If

'Reformat city and state
    temparry = Split(thisws.Range("B7"), ",")
    TempStr = StrConv(Trim(temparry(LBound(temparry))), vbProperCase)

Range("AA2") = TempStr & ", " & Trim(temparry(UBound(temparry)))


ElseIf review_type = 2 Then

'Client Information
Range("B2") = StrConv(thisws.Range("B2"), vbProperCase) 'client name
Range("C2") = Right(thisws.Range("AC10"), 2) & "/" & Left(thisws.Range("H10"), 9) 'county and case number

' Get county number and name in correct format
countynum = Right(thisws.Range("AC10"), 2) ' county
districtnum = thisws.Range("AH10") 'district
countylookup = countyformat(countynum, thisws.Range("O4"), districtnum) ' County Name
Range("AK2") = thisws.Range("O4") & " Assistance Office"
'Look for counties with districts
'If countynum = "02" Or countynum = "23" Or countynum = "40" Or _
'        countynum = "51" Or countynum = "63" Or countynum = "65" Then
'    Range("AK2") = thisws.Range("M5") & " Office"
'Else
'    Range("AK2") = thisws.Range("M5") & " CAO"
'End If

'Put into temp data file other data
Range("D2") = thisws.Range("a10") & thisws.Range("b10") & thisws.Range("c10") & thisws.Range("d10") & thisws.Range("e10") & thisws.Range("f10") 'Review Number
review_number = Range("D2")
Range("E2") = thisws.Range("AM10") 'Sample Month
sample_month = thisws.Range("AM10") 'Sample Month
Range("A7") = Val(thisws.Range("AO3") & thisws.Range("AP3")) 'Examiner Number
Range("G2") = Range("C7") & "/" & thisws.Range("AF2") 'Examiner Name
Range("E7") = review_type
'Range("E17") = Val(thisws.Range("K22"))
'reformat street address
If thisws.Range("B4") <> "" Then
    Range("Z2") = StrConv(thisws.Range("B3"), vbProperCase) & ", " & StrConv(thisws.Range("B4"), vbProperCase)
Else
    Range("Z2") = StrConv(thisws.Range("B3"), vbProperCase)
End If

'Reformat city and state
    temparry = Split(thisws.Range("B5"), ",")
    TempStr = StrConv(Trim(temparry(LBound(temparry))), vbProperCase)
    Range("AA2") = TempStr & ", " & Trim(temparry(UBound(temparry)))
    
Range("CE2") = thisws.Range("Q10") 'category
Range("CG2") = thisws.Range("Z10") 'grant group

Else

    MsgBox ("Preliminary Information Memo is not needed for this type of review")

End

End If

'Going out to website with list of Executive Directors
        FullTextFileName = datasourcews.Range("U6")
            
            'ActiveWorkbook.FollowHyperlink (FullTextFileName)
            Workbooks.Open fileName:=FullTextFileName, UpdateLinks:=False, ReadOnly:=True
            
            Range("A1:D1000").UnMerge

            'Allegheny Districts are formatted as dates, change them to 02/XX format
            For irow = 1 To 100
                If IsDate(Range("A" & irow)) Then
                    TempStr = WorksheetFunction.Text(Range("A" & irow), "mm/d/yyyy")
                    Range("A" & irow).NumberFormat = "@"
                    Range("A" & irow) = Left(TempStr, 4)
                End If
            Next irow
            
            'Look for county or district number
                found_row = Range("A1:A1000").Find(countylookup).Row
            
            datasourcews.Range("Y2") = Range("D" & found_row)
            'For irow = 1 To 1000
            '    If Range("A" & irow) = countylookup Then
                    'Get ED name
            '        datasourcews.Range("Y2") = Range("D" & irow)
            '        'Get CAO/DO Address
                    caoaddress = ""
                    For ii = 0 To 4
                        If Range("C" & found_row + ii) = "" Then
                            Exit For
                        End If
                        caoaddress = caoaddress & Range("C" & found_row + ii) & ", "
                    Next ii
                    datasourcews.Range("AL2") = caoaddress
                'End If
            'Next irow
            ActiveWorkbook.Close False
            
            'Application.Wait Now + TimeSerial(0, 0, 2)
            'CloseApp "CAO List", ""
            
        'put in deadline - 2 weeks
        datasourcews.Range("AQ2") = datasourcews.Range("A2") + 7
        datasourcewb.Close True
            
             If review_type = 2 Then
                sample_month = Left(sample_month, 2) & Right(sample_month, 4)
                findmemo = "MA Waiver Memo for Review Number " & review_number & " for Sample Month " & sample_month & ".doc"
             Else
                findmemo = "MA Waiver Memo for Review Number " & review_number & " for Sample Month " & sample_month & ".doc"
             End If

    'If review_type = 2 Then
    '    findmemo = "QC14 CAO Request Memo for Review Number " & review_number & " for Sample Month " & sample_month & ".docx"
    'Else
    '    findmemo = "QC14 Request Letter for Review Number " & review_number & " for Sample Month " & sample_month & ".doc"
    'End If
    'Copy template finding memo to active directory
    SrceFile = PathStr & "Finding Memo\MA Waiver Memo.docx"
    tempfindmemoname = sPath & "\FM temp.docx"
    DestFile = tempfindmemoname
    FileCopy SrceFile, DestFile

    Set appWD = CreateObject("Word.Application")
    appWD.Visible = True

    With appWD
    .Documents.Open fileName:=tempfindmemoname
    
    With .ActiveDocument.MailMerge
        .OpenDataSource Name:=(tempsourcename), _
            ReadOnly:=True, LinkToSource:=0, AddToRecentFiles:=False, _
            PasswordDocument:="", PasswordTemplate:="", WritePasswordDocument:="", _
            WritePasswordTemplate:="", Revert:=False, Format:=0, _
            Connection:="yourNamedRange", SQLStatement:="SELECT * FROM `yourNamedRange`", SQLStatement1:=""
        .Destination = 0
        .SuppressBlankLines = True
        .Execute Pause:=False
    End With
    
    '.ActiveDocument.Protect Type:=2, NoReset:=True
    
Application.StatusBar = "Now creating document"

    .ActiveDocument.Content.Font.Name = "Times New Roman"
    '.ActiveDocument.Content.Font.Size = 12

    .ActiveDocument.SaveAs sPath & "\" & findmemo
    .ActiveDocument.Close
    
    .Documents(tempfindmemoname).Close SaveChanges:=False

End With
appWD.Quit
Set appWD = Nothing

'delete temp files
Kill tempfindmemoname
Kill tempsourcename

Application.StatusBar = False
Application.ScreenUpdating = True
Application.Visible = True

MsgBox "A MA Waiver Memo has been saved as " & sPath & "\" & findmemo & _
". If needed, please close the Internet Explorer with the CAO Directory. - Thank you!"

End Sub
Sub MA_LTC_WAIVER()

'Create MA Waiver Memo
Dim appWD As Object
Dim review_number As Long
Dim sample_month As String, PathStr As String, findmemo As String
Dim countynum As String, districtnum As String

'Find path to Examiner's files on this PC
Set WshNetwork = CreateObject("WScript.Network")
Set oDrives = WshNetwork.EnumNetworkDrives

Dim DLetter As String, DUNC As String
Dim thisws As Worksheet, thiswb As Workbook
Dim sPath As String, tempsourcename As String

sPath = ActiveWorkbook.Path

Set thisws = ActiveSheet
Set thiswb = ActiveWorkbook

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

PathStr = DLetter
Application.ScreenUpdating = False

Application.StatusBar = "Starting mail merge ..."

'Copy template data source to active directory
tempsourcename = sPath & "\FM DS Temp.xlsx"
SrceFile = PathStr & "Finding Memo\Finding Memo Data Source.xlsx"
DestFile = tempsourcename
FileCopy SrceFile, DestFile

'Open spreadsheet where all the data is stored to populate findings memo
Workbooks.Open fileName:=tempsourcename, UpdateLinks:=False

Dim datasourcewb As Workbook, datasourcews As Worksheet

Set datasourcewb = ActiveWorkbook
Set datasourcews = ActiveSheet


'deleting information in Finding Memo Data Source spreadsheet
Range("B2:E2") = ""
Range("G2") = ""
Range("O2:S2") = ""
Range("W2:Y2") = ""

'determing what kind of review it is by the first number in the review
review_type = Left(thisws.Name, 1)

'SNAP Positive
If review_type = 5 Then

'Client Information
Range("B2") = StrConv(thisws.Range("B4"), vbProperCase)
Range("C2") = Right(thisws.Range("X18"), 2) & "/" & Left(thisws.Range("I18"), 9)

' Get county number and name in correct format
countynum = Right(thisws.Range("X18"), 2)
districtnum = thisws.Range("B153") & thisws.Range("C153")
countylookup = countyformat(countynum, thisws.Range("M5"), districtnum)
Range("AK2") = thisws.Range("M5") & " Assistance Office"
'Look for counties with districts
'If countynum = "02" Or countynum = "23" Or countynum = "40" Or _
'        countynum = "51" Or countynum = "63" Or countynum = "65" Then
'    Range("AK2") = thisws.Range("M5") & " Office"
'Else
'    Range("AK2") = thisws.Range("M5") & " CAO"
'End If

'Put into temp data file other data
Range("D2") = thisws.Range("A18")
review_number = Range("D2")
Range("E2") = thisws.Range("AD18") & "/" & thisws.Range("AG18")
sample_month = thisws.Range("AD18") & thisws.Range("AG18")
Range("A7") = Val(thisws.Range("AJ5") & thisws.Range("AK5"))
Range("G2") = Range("C7") & "/" & thisws.Range("AE5")
Range("E7") = review_type
Range("O2") = thiswb.Sheets("FS Computation").Range("B60")
Range("P2") = thiswb.Sheets("FS Computation").Range("C60")
Range("Q2") = thisws.Range("Y22")
Range("R2") = thisws.Range("B29") & " - " & thisws.Range("G29") & " - " & thisws.Range("K29")
Range("E17") = Val(thisws.Range("K22"))
'reformat street address
If thisws.Range("B6") <> "" Then
    Range("Z2") = StrConv(thisws.Range("B5"), vbProperCase) & ", " & StrConv(thisws.Range("B6"), vbProperCase)
Else
    Range("Z2") = StrConv(thisws.Range("B5"), vbProperCase)
End If

'Reformat city and state
    temparry = Split(thisws.Range("B7"), ",")
    TempStr = StrConv(Trim(temparry(LBound(temparry))), vbProperCase)

Range("AA2") = TempStr & ", " & Trim(temparry(UBound(temparry)))


ElseIf review_type = 2 Then

'Client Information
Range("B2") = StrConv(thisws.Range("B2"), vbProperCase) 'client name
Range("C2") = Right(thisws.Range("AC10"), 2) & "/" & Left(thisws.Range("H10"), 9) 'county and case number

' Get county number and name in correct format
countynum = Right(thisws.Range("AC10"), 2) ' county
districtnum = thisws.Range("AH10") 'district
countylookup = countyformat(countynum, thisws.Range("O4"), districtnum) ' County Name
Range("AK2") = thisws.Range("O4") & " Assistance Office"
'Look for counties with districts
'If countynum = "02" Or countynum = "23" Or countynum = "40" Or _
'        countynum = "51" Or countynum = "63" Or countynum = "65" Then
'    Range("AK2") = thisws.Range("M5") & " Office"
'Else
'    Range("AK2") = thisws.Range("M5") & " CAO"
'End If

'Put into temp data file other data
Range("D2") = thisws.Range("a10") & thisws.Range("b10") & thisws.Range("c10") & thisws.Range("d10") & thisws.Range("e10") & thisws.Range("f10") 'Review Number
review_number = Range("D2")
Range("E2") = thisws.Range("AM10") 'Sample Month
sample_month = thisws.Range("AM10") 'Sample Month
Range("A7") = Val(thisws.Range("AO3") & thisws.Range("AP3")) 'Examiner Number
Range("G2") = Range("C7") & "/" & thisws.Range("AF2") 'Examiner Name
Range("E7") = review_type
'Range("E17") = Val(thisws.Range("K22"))
'reformat street address
If thisws.Range("B4") <> "" Then
    Range("Z2") = StrConv(thisws.Range("B3"), vbProperCase) & ", " & StrConv(thisws.Range("B4"), vbProperCase)
Else
    Range("Z2") = StrConv(thisws.Range("B3"), vbProperCase)
End If

'Reformat city and state
    temparry = Split(thisws.Range("B5"), ",")
    TempStr = StrConv(Trim(temparry(LBound(temparry))), vbProperCase)
    Range("AA2") = TempStr & ", " & Trim(temparry(UBound(temparry)))
    
Range("CE2") = thisws.Range("Q10") 'category
Range("CG2") = thisws.Range("Z10") 'grant group

Else

    MsgBox ("Preliminary Information Memo is not needed for this type of review")

End

End If

'Going out to website with list of Executive Directors
        FullTextFileName = datasourcews.Range("U6")
            
            'ActiveWorkbook.FollowHyperlink (FullTextFileName)
            Workbooks.Open fileName:=FullTextFileName, UpdateLinks:=False, ReadOnly:=True
            
            Range("A1:D1000").UnMerge

            'Allegheny Districts are formatted as dates, change them to 02/XX format
            For irow = 1 To 100
                If IsDate(Range("A" & irow)) Then
                    TempStr = WorksheetFunction.Text(Range("A" & irow), "mm/d/yyyy")
                    Range("A" & irow).NumberFormat = "@"
                    Range("A" & irow) = Left(TempStr, 4)
                End If
            Next irow
            
            'Look for county or district number
                found_row = Range("A1:A1000").Find(countylookup).Row
            
            datasourcews.Range("Y2") = Range("D" & found_row)
            'For irow = 1 To 1000
            '    If Range("A" & irow) = countylookup Then
                    'Get ED name
            '        datasourcews.Range("Y2") = Range("D" & irow)
            '        'Get CAO/DO Address
                    caoaddress = ""
                    For ii = 0 To 4
                        If Range("C" & found_row + ii) = "" Then
                            Exit For
                        End If
                        caoaddress = caoaddress & Range("C" & found_row + ii) & ", "
                    Next ii
                    datasourcews.Range("AL2") = caoaddress
                'End If
            'Next irow
            ActiveWorkbook.Close False
            
            'Application.Wait Now + TimeSerial(0, 0, 2)
            'CloseApp "CAO List", ""
            
        'put in deadline - 2 weeks
        datasourcews.Range("AQ2") = datasourcews.Range("A2") + 7
        datasourcewb.Close True
            
             If review_type = 2 Then
                sample_month = Left(sample_month, 2) & Right(sample_month, 4)
                findmemo = "QC14 LTC Waiver Memo for Review Number " & review_number & " for Sample Month " & sample_month & ".doc"
             Else
                findmemo = "QC14 LTC Waiver Memo for Review Number " & review_number & " for Sample Month " & sample_month & ".doc"
             End If

    'If review_type = 2 Then
    '    findmemo = "QC14 CAO Request Memo for Review Number " & review_number & " for Sample Month " & sample_month & ".docx"
    'Else
    '    findmemo = "QC14 Request Letter for Review Number " & review_number & " for Sample Month " & sample_month & ".doc"
    'End If
    'Copy template finding memo to active directory
    SrceFile = PathStr & "Finding Memo\QC14 LTC-Waiver.docx"
    tempfindmemoname = sPath & "\FM temp.docx"
    DestFile = tempfindmemoname
    FileCopy SrceFile, DestFile

    Set appWD = CreateObject("Word.Application")
    appWD.Visible = True

    With appWD
    .Documents.Open fileName:=tempfindmemoname
    
    With .ActiveDocument.MailMerge
        .OpenDataSource Name:=(tempsourcename), _
            ReadOnly:=True, LinkToSource:=0, AddToRecentFiles:=False, _
            PasswordDocument:="", PasswordTemplate:="", WritePasswordDocument:="", _
            WritePasswordTemplate:="", Revert:=False, Format:=0, _
            Connection:="yourNamedRange", SQLStatement:="SELECT * FROM `yourNamedRange`", SQLStatement1:=""
        .Destination = 0
        .SuppressBlankLines = True
        .Execute Pause:=False
    End With
    
    '.ActiveDocument.Protect Type:=2, NoReset:=True
    
Application.StatusBar = "Now creating document"

    .ActiveDocument.Content.Font.Name = "Times New Roman"
    '.ActiveDocument.Content.Font.Size = 12

    .ActiveDocument.SaveAs sPath & "\" & findmemo
    .ActiveDocument.Close
    
    .Documents(tempfindmemoname).Close SaveChanges:=False

End With
appWD.Quit
Set appWD = Nothing

'delete temp files
Kill tempfindmemoname
Kill tempsourcename

Application.StatusBar = False
Application.ScreenUpdating = True
Application.Visible = True

MsgBox "A QC14 LTC Waiver Memo has been saved as " & sPath & "\" & findmemo & _
". If needed, please close the Internet Explorer with the CAO Directory. - Thank you!"

End Sub
Sub Time(IorF As String)

'Create Timeliness Form
Dim appWD As Object
Dim review_number As Long
Dim sample_month As String, PathStr As String, findmemo As String
Dim countynum As String, districtnum As String

'Find path to Examiner's files on this PC
Set WshNetwork = CreateObject("WScript.Network")
Set oDrives = WshNetwork.EnumNetworkDrives

Dim DLetter As String, DUNC As String
Dim thisws As Worksheet, thiswb As Workbook
Dim sPath As String, tempsourcename As String

sPath = ActiveWorkbook.Path

Set thisws = ActiveSheet
Set thiswb = ActiveWorkbook

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

PathStr = DLetter
Application.ScreenUpdating = False

Application.StatusBar = "Starting mail merge ..."

'Copy template data source to active directory
tempsourcename = sPath & "\FM DS Temp.xlsx"
SrceFile = PathStr & "Finding Memo\Finding Memo Data Source.xlsx"
DestFile = tempsourcename
FileCopy SrceFile, DestFile

'Open spreadsheet where all the data is stored to populate findings memo
Workbooks.Open fileName:=tempsourcename, UpdateLinks:=False

Dim datasourcewb As Workbook, datasourcews As Worksheet

Set datasourcewb = ActiveWorkbook
Set datasourcews = ActiveSheet


'deleting information in Finding Memo Data Source spreadsheet
Range("B2:E2") = ""
Range("G2") = ""
Range("O2:S2") = ""
Range("W2:Y2") = ""

'determing what kind of review it is by the first number in the review
review_type = Left(thisws.Name, 1)

'SNAP Positive
If review_type = 5 Then

'Client Information
Range("B2") = StrConv(thisws.Range("B4"), vbProperCase)
Range("C2") = Right(thisws.Range("X18"), 2) & "/" & Left(thisws.Range("I18"), 9)

' Get county number and name in correct format
countynum = Right(thisws.Range("X18"), 2)
districtnum = thisws.Range("B153") & thisws.Range("C153")
countylookup = countyformat(countynum, thisws.Range("M5"), districtnum)
Range("AK2") = thisws.Range("M5") & " Assistance Office"
'Look for counties with districts
'If countynum = "02" Or countynum = "23" Or countynum = "40" Or _
'        countynum = "51" Or countynum = "63" Or countynum = "65" Then
'    Range("AK2") = thisws.Range("M5") & " Office"
'Else
'    Range("AK2") = thisws.Range("M5") & " CAO"
'End If

'Put into temp data file other data
Range("D2") = thisws.Range("A18")
review_number = Range("D2")
Range("E2") = thisws.Range("AD18") & "/" & thisws.Range("AG18")
sample_month = thisws.Range("AD18") & thisws.Range("AG18")
Range("A7") = Val(thisws.Range("AJ5") & thisws.Range("AK5"))
Range("G2") = Range("C7") & "/" & thisws.Range("AE5")
Range("E7") = review_type
Range("O2") = thiswb.Sheets("FS Computation").Range("B60")
Range("P2") = thiswb.Sheets("FS Computation").Range("C60")
Range("Q2") = thisws.Range("Y22")
Range("R2") = thisws.Range("B29") & " - " & thisws.Range("G29") & " - " & thisws.Range("K29")
Range("E17") = Val(thisws.Range("K22"))
'reformat street address
If thisws.Range("B6") <> "" Then
    Range("Z2") = StrConv(thisws.Range("B5"), vbProperCase) & ", " & StrConv(thisws.Range("B6"), vbProperCase)
Else
    Range("Z2") = StrConv(thisws.Range("B5"), vbProperCase)
End If

'Reformat city and state
    temparry = Split(thisws.Range("B7"), ",")
    TempStr = StrConv(Trim(temparry(LBound(temparry))), vbProperCase)

Range("AA2") = TempStr & ", " & Trim(temparry(UBound(temparry)))



Else

    MsgBox ("SNAP Timeliness Memo not needed for this type of review")

End

End If

'Going out to website with list of Executive Directors
        FullTextFileName = datasourcews.Range("U6")
            
            'ActiveWorkbook.FollowHyperlink (FullTextFileName)
            Workbooks.Open fileName:=FullTextFileName, UpdateLinks:=False, ReadOnly:=True
            
            Range("A1:D1000").UnMerge

            'Allegheny Districts are formatted as dates, change them to 02/XX format
            For irow = 1 To 100
                If IsDate(Range("A" & irow)) Then
                    TempStr = WorksheetFunction.Text(Range("A" & irow), "mm/d/yyyy")
                    Range("A" & irow).NumberFormat = "@"
                    Range("A" & irow) = Left(TempStr, 4)
                End If
            Next irow
            
            'Look for county or district number
                found_row = Range("A1:A1000").Find(countylookup).Row
            
            datasourcews.Range("Y2") = Range("D" & found_row)
            'For irow = 1 To 1000
            '    If Range("A" & irow) = countylookup Then
                    'Get ED name
            '        datasourcews.Range("Y2") = Range("D" & irow)
            '        'Get CAO/DO Address
                    caoaddress = ""
                    For ii = 0 To 4
                        If Range("C" & found_row + ii) = "" Then
                            Exit For
                        End If
                        caoaddress = caoaddress & Range("C" & found_row + ii) & ", "
                    Next ii
                    datasourcews.Range("AL2") = caoaddress
                'End If
            'Next irow
            ActiveWorkbook.Close False
            
            'Application.Wait Now + TimeSerial(0, 0, 2)
            'CloseApp "CAO List", ""

            datasourcewb.Close True
            
    'Copy template finding memo to active directory
    If IorF = "Info" Then
        SrceFile = PathStr & "Finding Memo\Timeliness Information Memo.docx"
        findmemo = "SNAP Timeliness Information Memo for Review Number " & review_number & " for Sample Month " & sample_month & ".docx"
    Else
        SrceFile = PathStr & "Finding Memo\Timeliness Findings Memo.docx"
        findmemo = "SNAP Timeliness Findings Memo for Review Number " & review_number & " for Sample Month " & sample_month & ".docx"
    End If
    
    tempfindmemoname = sPath & "\FM temp.docx"
    DestFile = tempfindmemoname
    FileCopy SrceFile, DestFile

    Set appWD = CreateObject("Word.Application")
    appWD.Visible = True

    With appWD
    .Documents.Open fileName:=tempfindmemoname
    
    With .ActiveDocument.MailMerge
        .OpenDataSource Name:=(tempsourcename), _
            ReadOnly:=True, LinkToSource:=0, AddToRecentFiles:=False, _
            PasswordDocument:="", PasswordTemplate:="", WritePasswordDocument:="", _
            WritePasswordTemplate:="", Revert:=False, Format:=0, _
            Connection:="yourNamedRange", SQLStatement:="SELECT * FROM `yourNamedRange`", SQLStatement1:=""
        .Destination = 0
        .SuppressBlankLines = True
        .Execute Pause:=False
    End With
    
    '.ActiveDocument.Protect Type:=2, NoReset:=True
    
Application.StatusBar = "Now creating document"

    .ActiveDocument.Content.Font.Name = "Times New Roman"
    '.ActiveDocument.Content.Font.Size = 12

    .ActiveDocument.SaveAs sPath & "\" & findmemo
    .ActiveDocument.Close
    
    .Documents(tempfindmemoname).Close SaveChanges:=False

End With
appWD.Quit
Set appWD = Nothing

'delete temp files
Kill tempfindmemoname
Kill tempsourcename

Application.StatusBar = False
Application.ScreenUpdating = True
Application.Visible = True

MsgBox "A SNAP " & IorF & " Timeliness Memo has been saved as " & sPath & "\" & findmemo & _
". Please close the Internet Explorer with the CAO Directory. - Thank you!"

End Sub

Sub QC15()

'Create QC 15 Form
Dim appWD As Object
Dim review_number As Long
Dim sample_month As String, PathStr As String, findmemo As String
Dim countynum As String, districtnum As String

'Find path to Examiner's files on this PC
Set WshNetwork = CreateObject("WScript.Network")
Set oDrives = WshNetwork.EnumNetworkDrives

Dim DLetter As String, DUNC As String
Dim thisws As Worksheet, thiswb As Workbook
Dim sPath As String, tempsourcename As String

sPath = ActiveWorkbook.Path

Set thisws = ActiveSheet
Set thiswb = ActiveWorkbook

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

PathStr = DLetter
Application.ScreenUpdating = False

Application.StatusBar = "Starting mail merge ..."

'Copy template data source to active directory
tempsourcename = sPath & "\FM DS Temp.xlsx"
SrceFile = PathStr & "Finding Memo\Finding Memo Data Source.xlsx"
DestFile = tempsourcename
FileCopy SrceFile, DestFile

'Open spreadsheet where all the data is stored to populate findings memo
Workbooks.Open fileName:=tempsourcename, UpdateLinks:=False

Dim datasourcewb As Workbook, datasourcews As Worksheet

Set datasourcewb = ActiveWorkbook
Set datasourcews = ActiveSheet


'deleting information in Finding Memo Data Source spreadsheet
Range("B2:E2") = ""
Range("G2") = ""
Range("O2:S2") = ""
Range("W2:Y2") = ""

'determing what kind of review it is by the first number in the review
review_type = Left(thisws.Name, 1)

'MA
If review_type = 2 Then

'Client Information
Range("B2") = StrConv(thisws.Range("B2"), vbProperCase) 'client name
Range("C2") = Right(thisws.Range("AC10"), 2) & "/" & Left(thisws.Range("H10"), 9) 'county and case number

' Get county number and name in correct format
countynum = Right(thisws.Range("AC10"), 2) ' county
districtnum = thisws.Range("AH10") 'district
countylookup = countyformat(countynum, thisws.Range("O4"), districtnum) ' County Name
Range("AK2") = thisws.Range("O4") & " Assistance Office"
'Look for counties with districts
'If countynum = "02" Or countynum = "23" Or countynum = "40" Or _
'        countynum = "51" Or countynum = "63" Or countynum = "65" Then
'    Range("AK2") = thisws.Range("M5") & " Office"
'Else
'    Range("AK2") = thisws.Range("M5") & " CAO"
'End If

'Put into temp data file other data
Range("D2") = thisws.Range("a10") & thisws.Range("b10") & thisws.Range("c10") & thisws.Range("d10") & thisws.Range("e10") & thisws.Range("f10") 'Review Number
review_number = Range("D2")
Range("E2") = thisws.Range("AM10") 'Sample Month
sample_month = thisws.Range("AM10") 'Sample Month
Range("A7") = Val(thisws.Range("AO3") & thisws.Range("AP3")) 'Examiner Number
Range("G2") = Range("C7") & "/" & thisws.Range("AF2") 'Examiner Name
Range("E7") = review_type
'Range("E17") = Val(thisws.Range("K22"))
'reformat street address
If thisws.Range("B4") <> "" Then
    Range("Z2") = StrConv(thisws.Range("B3"), vbProperCase) & ", " & StrConv(thisws.Range("B4"), vbProperCase)
Else
    Range("Z2") = StrConv(thisws.Range("B3"), vbProperCase)
End If

'Reformat city and state
    temparry = Split(thisws.Range("B5"), ",")
    TempStr = StrConv(Trim(temparry(LBound(temparry))), vbProperCase)

Range("AA2") = TempStr & ", " & Trim(temparry(UBound(temparry)))

Else

    MsgBox ("QC15 Letter not needed for this type of review")

End

End If

'Going out to website with list of Executive Directors
        FullTextFileName = datasourcews.Range("U6")
            
            'ActiveWorkbook.FollowHyperlink (FullTextFileName)
            Workbooks.Open fileName:=FullTextFileName, UpdateLinks:=False, ReadOnly:=True
            
            Range("A1:D1000").UnMerge

            'Allegheny Districts are formatted as dates, change them to 02/XX format
            For irow = 1 To 100
                If IsDate(Range("A" & irow)) Then
                    TempStr = WorksheetFunction.Text(Range("A" & irow), "mm/d/yyyy")
                    Range("A" & irow).NumberFormat = "@"
                    Range("A" & irow) = Left(TempStr, 4)
                End If
            Next irow
            
            'Look for county or district number
                found_row = Range("A1:A1000").Find(countylookup).Row
            
            datasourcews.Range("Y2") = Range("D" & found_row)
            'For irow = 1 To 1000
            '    If Range("A" & irow) = countylookup Then
                    'Get ED name
            '        datasourcews.Range("Y2") = Range("D" & irow)
            '        'Get CAO/DO Address
                    caoaddress = ""
                    For ii = 0 To 4
                        If Range("C" & found_row + ii) = "" Then
                            Exit For
                        End If
                        caoaddress = caoaddress & Range("C" & found_row + ii) & ", "
                    Next ii
                    datasourcews.Range("AL2") = caoaddress
                'End If
            'Next irow
            ActiveWorkbook.Close False
            
            'Application.Wait Now + TimeSerial(0, 0, 2)
            'CloseApp "CAO List", ""

            datasourcewb.Close True
            
    'for MA schedule, take out the slash and keep the month and year numbers
    sample_month = Left(sample_month, 2) & Right(sample_month, 4)
            
    findmemo = "QC15 Letter for Review Number " & review_number & " for Sample Month " & sample_month & ".doc"

    'Copy template finding memo to active directory
    SrceFile = PathStr & "Finding Memo\MA QC15.docx"
    tempfindmemoname = sPath & "\FM temp.doc"
    DestFile = tempfindmemoname
    FileCopy SrceFile, DestFile

    Set appWD = CreateObject("Word.Application")
    appWD.Visible = True

    With appWD
    .Documents.Open fileName:=tempfindmemoname
    
    With .ActiveDocument.MailMerge
        .OpenDataSource Name:=(tempsourcename), _
            ReadOnly:=True, LinkToSource:=0, AddToRecentFiles:=False, _
            PasswordDocument:="", PasswordTemplate:="", WritePasswordDocument:="", _
            WritePasswordTemplate:="", Revert:=False, Format:=0, _
            Connection:="yourNamedRange", SQLStatement:="SELECT * FROM `yourNamedRange`", SQLStatement1:=""
        .Destination = 0
        .SuppressBlankLines = True
        .Execute Pause:=False
    End With
    
    '.ActiveDocument.Protect Type:=2, NoReset:=True
    
Application.StatusBar = "Now creating document"

    .ActiveDocument.Content.Font.Name = "Times New Roman"
    '.ActiveDocument.Content.Font.Size = 12

    .ActiveDocument.SaveAs sPath & "\" & findmemo
    .ActiveDocument.Close
    
    .Documents(tempfindmemoname).Close SaveChanges:=False

End With
appWD.Quit
Set appWD = Nothing

'delete temp files
Kill tempfindmemoname
Kill tempsourcename

Application.StatusBar = False
Application.ScreenUpdating = True
Application.Visible = True

MsgBox "A QC15 Letter has been saved as " & sPath & "\" & findmemo & _
". Please close the Internet Explorer with the CAO Directory. - Thank you!"

End Sub
Sub Post()

'Create Post Office Form
Dim appWD As Object
Dim review_number As Long
Dim sample_month As String, PathStr As String, findmemo As String
Dim countynum As String, districtnum As String

'Find path to Examiner's files on this PC
Set WshNetwork = CreateObject("WScript.Network")
Set oDrives = WshNetwork.EnumNetworkDrives

Dim DLetter As String, DUNC As String
Dim thisws As Worksheet, thiswb As Workbook
Dim sPath As String, tempsourcename As String

sPath = ActiveWorkbook.Path

Set thisws = ActiveSheet
Set thiswb = ActiveWorkbook

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

PathStr = DLetter
Application.ScreenUpdating = False

Application.StatusBar = "Starting mail merge ..."

'Copy template data source to active directory
tempsourcename = sPath & "\FM DS Temp.xlsx"
SrceFile = PathStr & "Finding Memo\Finding Memo Data Source.xlsx"
DestFile = tempsourcename
FileCopy SrceFile, DestFile

'Open spreadsheet where all the data is stored to populate findings memo
Workbooks.Open fileName:=tempsourcename, UpdateLinks:=False

Dim datasourcewb As Workbook, datasourcews As Worksheet

Set datasourcewb = ActiveWorkbook
Set datasourcews = ActiveSheet


'deleting information in Finding Memo Data Source spreadsheet
Range("B2:E2") = ""
Range("G2") = ""
Range("O2:S2") = ""
Range("W2:Y2") = ""

'determing what kind of review it is by the first number in the review
review_type = Left(thisws.Name, 1)

'SNAP Positive
If review_type = 5 Then

'Client Information
Range("B2") = StrConv(thisws.Range("B4"), vbProperCase)
Range("C2") = thisws.Range("C155") & thisws.Range("D155") & "/" & Left(thisws.Range("I18"), 9)

' Get county number and name in correct format
countynum = thisws.Range("C155") & thisws.Range("D155")
districtnum = thisws.Range("E155") & thisws.Range("F155")
countylookup = countyformat(countynum, thisws.Range("M5"), districtnum)
Range("AK2") = thisws.Range("M5") & " Assistance Office"

'Put into temp data file other data
Range("D2") = thisws.Range("A18")
review_number = Range("D2")
Range("E2") = thisws.Range("AD18") & "/" & thisws.Range("AG18")
sample_month = thisws.Range("AD18") & thisws.Range("AG18")
Range("A7") = Val(thisws.Range("AJ5") & thisws.Range("AK5"))
Range("G2") = Range("C7") & "/" & thisws.Range("AE5")
Range("E7") = review_type
Range("O2") = thiswb.Sheets("FS Computation").Range("B60")
Range("P2") = thiswb.Sheets("FS Computation").Range("C60")
Range("Q2") = thisws.Range("Y22")
Range("R2") = thisws.Range("B29") & " - " & thisws.Range("G29") & " - " & thisws.Range("K29")
Range("E17") = Val(thisws.Range("K22"))
'reformat street address
If thisws.Range("B6") <> "" Then
    Range("Z2") = StrConv(thisws.Range("B5"), vbProperCase) & ", " & StrConv(thisws.Range("B6"), vbProperCase)
Else
    Range("Z2") = StrConv(thisws.Range("B5"), vbProperCase)
End If

'Reformat city and state
    temparry = Split(thisws.Range("B7"), ",")
    TempStr = StrConv(Trim(temparry(LBound(temparry))), vbProperCase)

Range("AA2") = TempStr & ", " & Trim(temparry(UBound(temparry)))

End If

    datasourcewb.Close True
    findmemo = "Post Office Letter for Review Number " & review_number & " for Sample Month " & sample_month & ".docx"

    'Copy template finding memo to active directory
    SrceFile = PathStr & "Finding Memo\Post Office.docx"
    tempfindmemoname = sPath & "\FM temp.doc"
    DestFile = tempfindmemoname
    FileCopy SrceFile, DestFile

    Set appWD = CreateObject("Word.Application")
    appWD.Visible = True

    With appWD
    .Documents.Open fileName:=tempfindmemoname
    
    With .ActiveDocument.MailMerge
        .OpenDataSource Name:=(tempsourcename), _
            ReadOnly:=True, LinkToSource:=0, AddToRecentFiles:=False, _
            PasswordDocument:="", PasswordTemplate:="", WritePasswordDocument:="", _
            WritePasswordTemplate:="", Revert:=False, Format:=0, _
            Connection:="yourNamedRange", SQLStatement:="SELECT * FROM `yourNamedRange`", SQLStatement1:=""
        .Destination = 0
        .SuppressBlankLines = True
        .Execute Pause:=False
    End With
    
Application.StatusBar = "Now creating document"

    .ActiveDocument.Content.Font.Name = "Times New Roman"
    '.ActiveDocument.Content.Font.Size = 12

    .ActiveDocument.SaveAs sPath & "\" & findmemo
    .ActiveDocument.Close
    
    .Documents(tempfindmemoname).Close SaveChanges:=False

End With
appWD.Quit
Set appWD = Nothing

'delete temp files
Kill tempfindmemoname
Kill tempsourcename

Application.StatusBar = False
Application.ScreenUpdating = True
Application.Visible = True

MsgBox "A Post Office Letter has been saved as " & sPath & "\" & findmemo & _
".  Thank you!"

End Sub

Sub SNAPPend()

'Create Post Office Form
Dim appWD As Object
Dim review_number As Long
Dim sample_month As String, PathStr As String, findmemo As String
Dim countynum As String, districtnum As String

'Find path to Examiner's files on this PC
Set WshNetwork = CreateObject("WScript.Network")
Set oDrives = WshNetwork.EnumNetworkDrives

Dim DLetter As String, DUNC As String
Dim thisws As Worksheet, thiswb As Workbook
Dim sPath As String, tempsourcename As String

sPath = ActiveWorkbook.Path

Set thisws = ActiveSheet
Set thiswb = ActiveWorkbook

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

PathStr = DLetter
Application.ScreenUpdating = False

Application.StatusBar = "Starting mail merge ..."

'Copy template data source to active directory
tempsourcename = sPath & "\FM DS Temp.xlsx"
SrceFile = PathStr & "Finding Memo\Finding Memo Data Source.xlsx"
DestFile = tempsourcename
FileCopy SrceFile, DestFile

'Open spreadsheet where all the data is stored to populate findings memo
Workbooks.Open fileName:=tempsourcename, UpdateLinks:=False

Dim datasourcewb As Workbook, datasourcews As Worksheet

Set datasourcewb = ActiveWorkbook
Set datasourcews = ActiveSheet


'deleting information in Finding Memo Data Source spreadsheet
Range("B2:E2") = ""
Range("G2") = ""
Range("O2:S2") = ""
Range("W2:Y2") = ""

'determing what kind of review it is by the first number in the review
review_type = Left(thisws.Name, 1)

'SNAP Positive
If review_type = 5 Then

'Client Information
Range("B2") = StrConv(thisws.Range("B4"), vbProperCase)
Range("C2") = thisws.Range("C155") & thisws.Range("D155") & "/" & Left(thisws.Range("I18"), 9)

' Get county number and name in correct format
countynum = thisws.Range("C155") & thisws.Range("D155")
districtnum = thisws.Range("E155") & thisws.Range("F155")
countylookup = countyformat(countynum, thisws.Range("M5"), districtnum)
Range("AK2") = thisws.Range("M5") & " Assistance Office"

'Put into temp data file other data
Range("D2") = thisws.Range("A18")
review_number = Range("D2")
Range("E2") = thisws.Range("AD18") & "/" & thisws.Range("AG18")
sample_month = thisws.Range("AD18") & thisws.Range("AG18")
Range("A7") = Val(thisws.Range("AJ5") & thisws.Range("AK5"))
Range("G2") = Range("C7") & "/" & thisws.Range("AE5")
Range("E7") = review_type
Range("O2") = thiswb.Sheets("FS Computation").Range("B60")
Range("P2") = thiswb.Sheets("FS Computation").Range("C60")
Range("Q2") = thisws.Range("Y22")
Range("R2") = thisws.Range("B29") & " - " & thisws.Range("G29") & " - " & thisws.Range("K29")
Range("E17") = Val(thisws.Range("K22"))
'reformat street address
If thisws.Range("B6") <> "" Then
    Range("Z2") = StrConv(thisws.Range("B5"), vbProperCase) & ", " & StrConv(thisws.Range("B6"), vbProperCase)
Else
    Range("Z2") = StrConv(thisws.Range("B5"), vbProperCase)
End If

'Reformat city and state
    temparry = Split(thisws.Range("B7"), ",")
    TempStr = StrConv(Trim(temparry(LBound(temparry))), vbProperCase)

Range("AA2") = TempStr & ", " & Trim(temparry(UBound(temparry)))

End If

    datasourcewb.Close True
    findmemo = "SNAP Pending Letter for Review Number " & review_number & " for Sample Month " & sample_month & ".docx"

    'Copy template finding memo to active directory
    SrceFile = PathStr & "Finding Memo\SNAP-Pending Letter.docx"
    tempfindmemoname = sPath & "\FM temp.doc"
    DestFile = tempfindmemoname
    FileCopy SrceFile, DestFile

    Set appWD = CreateObject("Word.Application")
    appWD.Visible = True

    With appWD
    .Documents.Open fileName:=tempfindmemoname
    
    With .ActiveDocument.MailMerge
        .OpenDataSource Name:=(tempsourcename), _
            ReadOnly:=True, LinkToSource:=0, AddToRecentFiles:=False, _
            PasswordDocument:="", PasswordTemplate:="", WritePasswordDocument:="", _
            WritePasswordTemplate:="", Revert:=False, Format:=0, _
            Connection:="yourNamedRange", SQLStatement:="SELECT * FROM `yourNamedRange`", SQLStatement1:=""
        .Destination = 0
        .SuppressBlankLines = True
        .Execute Pause:=False
    End With
    
Application.StatusBar = "Now creating document"

    .ActiveDocument.Content.Font.Name = "Times New Roman"
    '.ActiveDocument.Content.Font.Size = 12

    .ActiveDocument.SaveAs sPath & "\" & findmemo
    .ActiveDocument.Close
    
    .Documents(tempfindmemoname).Close SaveChanges:=False

End With
appWD.Quit
Set appWD = Nothing

'delete temp files
Kill tempfindmemoname
Kill tempsourcename

Application.StatusBar = False
Application.ScreenUpdating = True
Application.Visible = True

MsgBox "A SNAP Pending Letter has been saved as " & sPath & "\" & findmemo & _
".  Thank you!"

End Sub

Sub SNAPDrop()

'Create Post Office Form
Dim appWD As Object
Dim review_number As Long
Dim sample_month As String, PathStr As String, findmemo As String
Dim countynum As String, districtnum As String

'Find path to Examiner's files on this PC
Set WshNetwork = CreateObject("WScript.Network")
Set oDrives = WshNetwork.EnumNetworkDrives

Dim DLetter As String, DUNC As String
Dim thisws As Worksheet, thiswb As Workbook
Dim sPath As String, tempsourcename As String

sPath = ActiveWorkbook.Path

Set thisws = ActiveSheet
Set thiswb = ActiveWorkbook

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

PathStr = DLetter
Application.ScreenUpdating = False

Application.StatusBar = "Starting mail merge ..."

'Copy template data source to active directory
tempsourcename = sPath & "\FM DS Temp.xlsx"
SrceFile = PathStr & "Finding Memo\Finding Memo Data Source.xlsx"
DestFile = tempsourcename
FileCopy SrceFile, DestFile

'Open spreadsheet where all the data is stored to populate findings memo
Workbooks.Open fileName:=tempsourcename, UpdateLinks:=False

Dim datasourcewb As Workbook, datasourcews As Worksheet

Set datasourcewb = ActiveWorkbook
Set datasourcews = ActiveSheet


'deleting information in Finding Memo Data Source spreadsheet
Range("B2:E2") = ""
Range("G2") = ""
Range("O2:S2") = ""
Range("W2:Y2") = ""

'determing what kind of review it is by the first number in the review
review_type = Left(thisws.Name, 1)

'SNAP Positive
If review_type = 5 Then

'Client Information
Range("B2") = StrConv(thisws.Range("B4"), vbProperCase)
Range("C2") = thisws.Range("C155") & thisws.Range("D155") & "/" & Left(thisws.Range("I18"), 9)

' Get county number and name in correct format
countynum = thisws.Range("C155") & thisws.Range("D155")
districtnum = thisws.Range("E155") & thisws.Range("F155")
countylookup = countyformat(countynum, thisws.Range("M5"), districtnum)
Range("AK2") = thisws.Range("M5") & " Assistance Office"

'Put into temp data file other data
Range("D2") = thisws.Range("A18")
review_number = Range("D2")
Range("E2") = thisws.Range("AD18") & "/" & thisws.Range("AG18")
sample_month = thisws.Range("AD18") & thisws.Range("AG18")
Range("A7") = Val(thisws.Range("AJ5") & thisws.Range("AK5"))
Range("G2") = Range("C7") & "/" & thisws.Range("AE5")
Range("E7") = review_type
Range("O2") = thiswb.Sheets("FS Computation").Range("B60")
Range("P2") = thiswb.Sheets("FS Computation").Range("C60")
Range("Q2") = thisws.Range("Y22")
Range("R2") = thisws.Range("B29") & " - " & thisws.Range("G29") & " - " & thisws.Range("K29")
Range("E17") = Val(thisws.Range("K22"))
Range("DC2") = thisws.Range("C22")
'reformat street address
If thisws.Range("B6") <> "" Then
    Range("Z2") = StrConv(thisws.Range("B5"), vbProperCase) & ", " & StrConv(thisws.Range("B6"), vbProperCase)
Else
    Range("Z2") = StrConv(thisws.Range("B5"), vbProperCase)
End If

'Reformat city and state
    temparry = Split(thisws.Range("B7"), ",")
    TempStr = StrConv(Trim(temparry(LBound(temparry))), vbProperCase)

Range("AA2") = TempStr & ", " & Trim(temparry(UBound(temparry)))

End If

    datasourcewb.Close True
    findmemo = "SNAP Drop Worksheet for Review Number " & review_number & " for Sample Month " & sample_month & ".docx"

    'Copy template finding memo to active directory
    SrceFile = PathStr & "Finding Memo\SNAP Active Drop Review.docx"
    tempfindmemoname = sPath & "\FM temp.doc"
    DestFile = tempfindmemoname
    FileCopy SrceFile, DestFile

    Set appWD = CreateObject("Word.Application")
    appWD.Visible = True

    With appWD
    .Documents.Open fileName:=tempfindmemoname
    
    With .ActiveDocument.MailMerge
        .OpenDataSource Name:=(tempsourcename), _
            ReadOnly:=True, LinkToSource:=0, AddToRecentFiles:=False, _
            PasswordDocument:="", PasswordTemplate:="", WritePasswordDocument:="", _
            WritePasswordTemplate:="", Revert:=False, Format:=0, _
            Connection:="yourNamedRange", SQLStatement:="SELECT * FROM `yourNamedRange`", SQLStatement1:=""
        .Destination = 0
        .SuppressBlankLines = True
        .Execute Pause:=False
    End With
    
Application.StatusBar = "Now creating document"

    .ActiveDocument.Content.Font.Name = "Times New Roman"
    '.ActiveDocument.Content.Font.Size = 12

    .ActiveDocument.SaveAs sPath & "\" & findmemo
    .ActiveDocument.Close
    
    .Documents(tempfindmemoname).Close SaveChanges:=False

End With
appWD.Quit
Set appWD = Nothing

'delete temp files
Kill tempfindmemoname
Kill tempsourcename

Application.StatusBar = False
Application.ScreenUpdating = True
Application.Visible = True

MsgBox "A SNAP Drop Worksheet has been saved as " & sPath & "\" & findmemo & _
".  Thank you!"

End Sub


Sub SNAPCAORequest()

'Create Post Office Form
Dim appWD As Object
Dim review_number As Long
Dim sample_month As String, PathStr As String, findmemo As String
Dim countynum As String, districtnum As String

'Find path to Examiner's files on this PC
Set WshNetwork = CreateObject("WScript.Network")
Set oDrives = WshNetwork.EnumNetworkDrives

Dim DLetter As String, DUNC As String
Dim thisws As Worksheet, thiswb As Workbook
Dim sPath As String, tempsourcename As String

sPath = ActiveWorkbook.Path

Set thisws = ActiveSheet
Set thiswb = ActiveWorkbook

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

PathStr = DLetter
Application.ScreenUpdating = False

Application.StatusBar = "Starting mail merge ..."

'Copy template data source to active directory
tempsourcename = sPath & "\FM DS Temp.xlsx"
SrceFile = PathStr & "Finding Memo\Finding Memo Data Source.xlsx"
DestFile = tempsourcename
FileCopy SrceFile, DestFile

'Open spreadsheet where all the data is stored to populate findings memo
Workbooks.Open fileName:=tempsourcename, UpdateLinks:=False

Dim datasourcewb As Workbook, datasourcews As Worksheet

Set datasourcewb = ActiveWorkbook
Set datasourcews = ActiveSheet


'deleting information in Finding Memo Data Source spreadsheet
Range("B2:E2") = ""
Range("G2") = ""
Range("O2:S2") = ""
Range("W2:Y2") = ""

'determing what kind of review it is by the first number in the review
review_type = Left(thisws.Name, 1)

'SNAP Positive
If review_type = 5 Then

'Client Information
Range("B2") = StrConv(thisws.Range("B4"), vbProperCase)
Range("C2") = thisws.Range("C155") & thisws.Range("D155") & "/" & Left(thisws.Range("I18"), 9)

' Get county number and name in correct format
countynum = thisws.Range("C155") & thisws.Range("D155")
districtnum = thisws.Range("E155") & thisws.Range("F155")
countylookup = countyformat(countynum, thisws.Range("M5"), districtnum)
Range("AK2") = thisws.Range("M5") & " Assistance Office"

'Put into temp data file other data
Range("D2") = thisws.Range("A18")
review_number = Range("D2")
Range("E2") = thisws.Range("AD18") & "/" & thisws.Range("AG18")
sample_month = thisws.Range("AD18") & thisws.Range("AG18")
Range("A7") = Val(thisws.Range("AJ5") & thisws.Range("AK5"))
Range("G2") = Range("C7") & "/" & thisws.Range("AE5")
Range("E7") = review_type
Range("O2") = thiswb.Sheets("FS Computation").Range("B60")
Range("P2") = thiswb.Sheets("FS Computation").Range("C60")
Range("Q2") = thisws.Range("Y22")
Range("R2") = thisws.Range("B29") & " - " & thisws.Range("G29") & " - " & thisws.Range("K29")
Range("E17") = Val(thisws.Range("K22"))
Range("DC2") = thisws.Range("C22")
'reformat street address
If thisws.Range("B6") <> "" Then
    Range("Z2") = StrConv(thisws.Range("B5"), vbProperCase) & ", " & StrConv(thisws.Range("B6"), vbProperCase)
Else
    Range("Z2") = StrConv(thisws.Range("B5"), vbProperCase)
End If

'Reformat city and state
    temparry = Split(thisws.Range("B7"), ",")
    TempStr = StrConv(Trim(temparry(LBound(temparry))), vbProperCase)

Range("AA2") = TempStr & ", " & Trim(temparry(UBound(temparry)))

'SNAP Negative
ElseIf review_type = 6 Then

'Client Information
Range("B2") = StrConv(thisws.Range("C8"), vbProperCase)
Range("C2") = Right(thisws.Range("AA20"), 2) & "/" & Right(thisws.Range("L20"), 8)

' Get county number and name in correct format
countynum = Right(thisws.Range("AA20"), 2)
districtnum = thisws.Range("H56") & thisws.Range("I56")
countylookup = countyformat(countynum, thisws.Range("C1"), districtnum)
Range("AK2") = thisws.Range("M5") & " Assistance Office"


End If

    datasourcewb.Close True
    findmemo = "SNAP CAO Request Form for Review Number " & review_number & " for Sample Month " & sample_month & ".docx"

    'Copy template finding memo to active directory
    SrceFile = PathStr & "Finding Memo\SNAP Support Forms.docx"
    tempfindmemoname = sPath & "\FM temp.doc"
    DestFile = tempfindmemoname
    FileCopy SrceFile, DestFile

    Set appWD = CreateObject("Word.Application")
    appWD.Visible = True

    With appWD
    .Documents.Open fileName:=tempfindmemoname
    
    With .ActiveDocument.MailMerge
        .OpenDataSource Name:=(tempsourcename), _
            ReadOnly:=True, LinkToSource:=0, AddToRecentFiles:=False, _
            PasswordDocument:="", PasswordTemplate:="", WritePasswordDocument:="", _
            WritePasswordTemplate:="", Revert:=False, Format:=0, _
            Connection:="yourNamedRange", SQLStatement:="SELECT * FROM `yourNamedRange`", SQLStatement1:=""
        .Destination = 0
        .SuppressBlankLines = True
        .Execute Pause:=False
    End With
    
Application.StatusBar = "Now creating document"

    .ActiveDocument.Content.Font.Name = "Times New Roman"
    '.ActiveDocument.Content.Font.Size = 12

    .ActiveDocument.SaveAs sPath & "\" & findmemo
    .ActiveDocument.Close
    
    .Documents(tempfindmemoname).Close SaveChanges:=False

End With
appWD.Quit
Set appWD = Nothing

'delete temp files
Kill tempfindmemoname
Kill tempsourcename

Application.StatusBar = False
Application.ScreenUpdating = True
Application.Visible = True

MsgBox "A SNAP CAO Request Form has been saved as " & sPath & "\" & findmemo & _
".  Thank you!"

End Sub

Sub TANFCAORequest()

'Create Post Office Form
Dim appWD As Object
Dim review_number As Long
Dim sample_month As String, PathStr As String, findmemo As String
Dim countynum As String, districtnum As String

'Find path to Examiner's files on this PC
Set WshNetwork = CreateObject("WScript.Network")
Set oDrives = WshNetwork.EnumNetworkDrives

Dim DLetter As String, DUNC As String
Dim thisws As Worksheet, thiswb As Workbook
Dim sPath As String, tempsourcename As String

sPath = ActiveWorkbook.Path

Set thisws = ActiveSheet
Set thiswb = ActiveWorkbook

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

PathStr = DLetter
Application.ScreenUpdating = False

Application.StatusBar = "Starting mail merge ..."

'Copy template data source to active directory
tempsourcename = sPath & "\FM DS Temp.xlsx"
SrceFile = PathStr & "Finding Memo\Finding Memo Data Source.xlsx"
DestFile = tempsourcename
FileCopy SrceFile, DestFile

'Open spreadsheet where all the data is stored to populate findings memo
Workbooks.Open fileName:=tempsourcename, UpdateLinks:=False

Dim datasourcewb As Workbook, datasourcews As Worksheet

Set datasourcewb = ActiveWorkbook
Set datasourcews = ActiveSheet


'deleting information in Finding Memo Data Source spreadsheet
Range("B2:E2") = ""
Range("G2") = ""
Range("O2:S2") = ""
Range("W2:Y2") = ""

'determing what kind of review it is by the first number in the review
review_type = Left(thisws.Name, 1)

'TANF
'If review_type = 1 Then
Select Case review_type
Case Is = 1, 9
'Client Information
Range("B2") = StrConv(thisws.Range("B2"), vbProperCase)
Range("C2") = Right(thisws.Range("U10"), 2) & "/" & Left(thisws.Range("I10"), 9)

' Get county number and name in correct format
countynum = Right(thisws.Range("U10"), 2)
districtnum = thisws.Range("Y10")
countylookup = countyformat(countynum, thisws.Range("O4"), districtnum)
Range("AK2") = thisws.Range("O4") & " Assistance Office"

'Put into temp data file other data
Range("D2") = thisws.Range("A10")
review_number = Range("D2")
Range("E2") = thisws.Range("AB10")
sample_month = thisws.Range("AB10")
Range("A7") = Val(thisws.Range("AO3") & thisws.Range("AP3"))
Range("G2") = Range("C7") & "/" & thisws.Range("AG4")
Range("E7") = review_type
Range("DC2") = thisws.Range("AI10")
'reformat street address
    'If thisws.Range("B5") <> "" Then
    '    datasourcews.Range("A6") = StrConv(thisws.Range("B3"), vbProperCase) & ", " & StrConv(thisws.Range("B4"), vbProperCase)
    'Else
        datasourcews.Range("A6") = StrConv(thisws.Range("B3"), vbProperCase)
    'End If
'Reformat city and state
If review_type = 1 Then
    temparry = Split(thisws.Range("B4"), ",")
    TempStr = StrConv(Trim(temparry(LBound(temparry))), vbProperCase)
    datasourcews.Range("A7") = TempStr & ", " & Trim(temparry(UBound(temparry)))
Else
    temparry = Split(thisws.Range("B5"), ",")
    TempStr = StrConv(Trim(temparry(LBound(temparry))), vbProperCase)
    datasourcews.Range("A7") = TempStr & ", " & Trim(temparry(UBound(temparry)))
End If

Range("AA2") = TempStr & ", " & Trim(temparry(UBound(temparry)))
End Select
'End If

    datasourcewb.Close True
    If review_type = 1 Then
    findmemo = "TANF CAO Request Form for Review Number " & review_number & " for Sample Month " & sample_month & ".docx"
    Else
    findmemo = "GA CAO Request Form for Review Number " & review_number & " for Sample Month " & sample_month & ".docx"
    End If

    'Copy template finding memo to active directory
    SrceFile = PathStr & "Finding Memo\QC CR Matls Request Form.docx"
    tempfindmemoname = sPath & "\FM temp.doc"
    DestFile = tempfindmemoname
    FileCopy SrceFile, DestFile

    Set appWD = CreateObject("Word.Application")
    appWD.Visible = True

    With appWD
    .Documents.Open fileName:=tempfindmemoname
    
    With .ActiveDocument.MailMerge
        .OpenDataSource Name:=(tempsourcename), _
            ReadOnly:=True, LinkToSource:=0, AddToRecentFiles:=False, _
            PasswordDocument:="", PasswordTemplate:="", WritePasswordDocument:="", _
            WritePasswordTemplate:="", Revert:=False, Format:=0, _
            Connection:="yourNamedRange", SQLStatement:="SELECT * FROM `yourNamedRange`", SQLStatement1:=""
        .Destination = 0
        .SuppressBlankLines = True
        .Execute Pause:=False
    End With
    
Application.StatusBar = "Now creating document"

    .ActiveDocument.Content.Font.Name = "Times New Roman"
    '.ActiveDocument.Content.Font.Size = 12

    .ActiveDocument.SaveAs sPath & "\" & findmemo
    .ActiveDocument.Close
    
    .Documents(tempfindmemoname).Close SaveChanges:=False

End With
appWD.Quit
Set appWD = Nothing

'delete temp files
Kill tempfindmemoname
Kill tempsourcename

Application.StatusBar = False
Application.ScreenUpdating = True
Application.Visible = True
If review_type = 1 Then
MsgBox "A TANF CAO Request Form has been saved as " & sPath & "\" & findmemo & _
".  Thank you!"
Else
MsgBox "A GA CAO Request Form has been saved as " & sPath & "\" & findmemo & _
".  Thank you!"
End If

End Sub
Sub SpPend()

'Create Post Office Form
Dim appWD As Object
Dim review_number As Long
Dim sample_month As String, PathStr As String, findmemo As String
Dim countynum As String, districtnum As String

'Find path to Examiner's files on this PC
Set WshNetwork = CreateObject("WScript.Network")
Set oDrives = WshNetwork.EnumNetworkDrives

Dim DLetter As String, DUNC As String
Dim thisws As Worksheet, thiswb As Workbook
Dim sPath As String, tempsourcename As String

sPath = ActiveWorkbook.Path

Set thisws = ActiveSheet
Set thiswb = ActiveWorkbook

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

PathStr = DLetter
Application.ScreenUpdating = False

Application.StatusBar = "Starting mail merge ..."

'Copy template data source to active directory
tempsourcename = sPath & "\FM DS Temp.xlsx"
SrceFile = PathStr & "Finding Memo\Finding Memo Data Source.xlsx"
DestFile = tempsourcename
FileCopy SrceFile, DestFile

'Open spreadsheet where all the data is stored to populate findings memo
Workbooks.Open fileName:=tempsourcename, UpdateLinks:=False

Dim datasourcewb As Workbook, datasourcews As Worksheet

Set datasourcewb = ActiveWorkbook
Set datasourcews = ActiveSheet


'deleting information in Finding Memo Data Source spreadsheet
Range("B2:E2") = ""
Range("G2") = ""
Range("O2:S2") = ""
Range("W2:Y2") = ""

'determing what kind of review it is by the first number in the review
review_type = Left(thisws.Name, 1)

'SNAP Positive
If review_type = 5 Then

'Client Information
Range("B2") = StrConv(thisws.Range("B4"), vbProperCase)
Range("C2") = thisws.Range("C155") & thisws.Range("D155") & "/" & Left(thisws.Range("I18"), 9)

' Get county number and name in correct format
countynum = thisws.Range("C155") & thisws.Range("D155")
districtnum = thisws.Range("E155") & thisws.Range("F155")
countylookup = countyformat(countynum, thisws.Range("M5"), districtnum)
Range("AK2") = thisws.Range("M5") & " Assistance Office"

'Put into temp data file other data
Range("D2") = thisws.Range("A18")
review_number = Range("D2")
Range("E2") = thisws.Range("AD18") & "/" & thisws.Range("AG18")
sample_month = thisws.Range("AD18") & thisws.Range("AG18")
Range("A7") = Val(thisws.Range("AJ5") & thisws.Range("AK5"))
Range("G2") = Range("C7") & "/" & thisws.Range("AE5")
Range("E7") = review_type
Range("O2") = thiswb.Sheets("FS Computation").Range("B60")
Range("P2") = thiswb.Sheets("FS Computation").Range("C60")
Range("Q2") = thisws.Range("Y22")
Range("R2") = thisws.Range("B29") & " - " & thisws.Range("G29") & " - " & thisws.Range("K29")
Range("E17") = Val(thisws.Range("K22"))
'reformat street address
If thisws.Range("B6") <> "" Then
    Range("Z2") = StrConv(thisws.Range("B5"), vbProperCase) & ", " & StrConv(thisws.Range("B6"), vbProperCase)
Else
    Range("Z2") = StrConv(thisws.Range("B5"), vbProperCase)
End If

'Reformat city and state
    temparry = Split(thisws.Range("B7"), ",")
    TempStr = StrConv(Trim(temparry(LBound(temparry))), vbProperCase)

Range("AA2") = TempStr & ", " & Trim(temparry(UBound(temparry)))

End If

    datasourcewb.Close True
    findmemo = "SNAP Spanish Pending Letter for Review Number " & review_number & " for Sample Month " & sample_month & ".docx"

    'Copy template finding memo to active directory
    SrceFile = PathStr & "Finding Memo\SNAP Pending Letter Spanish.docx"
    tempfindmemoname = sPath & "\FM temp.doc"
    DestFile = tempfindmemoname
    FileCopy SrceFile, DestFile

    Set appWD = CreateObject("Word.Application")
    appWD.Visible = True

    With appWD
    .Documents.Open fileName:=tempfindmemoname
    
    With .ActiveDocument.MailMerge
        .OpenDataSource Name:=(tempsourcename), _
            ReadOnly:=True, LinkToSource:=0, AddToRecentFiles:=False, _
            PasswordDocument:="", PasswordTemplate:="", WritePasswordDocument:="", _
            WritePasswordTemplate:="", Revert:=False, Format:=0, _
            Connection:="yourNamedRange", SQLStatement:="SELECT * FROM `yourNamedRange`", SQLStatement1:=""
        .Destination = 0
        .SuppressBlankLines = True
        .Execute Pause:=False
    End With
    
Application.StatusBar = "Now creating document"

    .ActiveDocument.Content.Font.Name = "Times New Roman"
    '.ActiveDocument.Content.Font.Size = 12

    .ActiveDocument.SaveAs sPath & "\" & findmemo
    .ActiveDocument.Close
    
    .Documents(tempfindmemoname).Close SaveChanges:=False

End With
appWD.Quit
Set appWD = Nothing

'delete temp files
Kill tempfindmemoname
Kill tempsourcename

Application.StatusBar = False
Application.ScreenUpdating = True
Application.Visible = True

MsgBox "A SNAP Spanish Pending Letter has been saved as " & sPath & "\" & findmemo & _
".  Thank you!"

End Sub
