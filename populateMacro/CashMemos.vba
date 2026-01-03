Attribute VB_Name = "CashMemos"
Sub AMR()

'Create Finding Memo
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

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'TANF
If review_type = 1 Then

' Get county number and name in correct format
countynum = Right(thisws.Range("U10"), 2)
districtnum = thisws.Range("Y10")
countylookup = countyformat(countynum, thisws.Range("O4"), districtnum)
Range("AK2") = thisws.Range("O4") & " Assistance Office"
'Look for counties with districts
'If countynum = "02" Or countynum = "23" Or countynum = "40" Or _
'        countynum = "51" Or countynum = "63" Or countynum = "65" Then
'    Range("AK2") = thisws.Range("O4") & " Office"
'Else
'    Range("AK2") = thisws.Range("O4") & " CAO"
'End If

'Client Information
Range("B2") = StrConv(thisws.Range("B2"), vbProperCase) 'Client Name
Range("C2") = Right(thisws.Range("U10"), 2) & "/" & thisws.Range("I10") 'County and Case Number
Range("D2") = thisws.Range("A10") 'Review Number
review_number = Range("D2")
Range("E2") = Left(thisws.Range("AB10"), 2) & "/" & Right(thisws.Range("AB10"), 4) 'Review Month
sample_month = thisws.Range("AB10")
Range("A7") = Val(thisws.Range("AO3") & thisws.Range("AP3")) 'Reviewer Number
Range("G2") = Range("C7") & "/" & thisws.Range("AG2") 'Supervisor and Reviewer Name
Range("E7") = review_type
Range("F7") = review_type
Range("R2") = thisws.Range("J61") & " - " & thisws.Range("O61") & " - " & thisws.Range("T61") 'Element, Nature, & Cause Code
Range("E17") = Val(thisws.Range("AL10")) 'Review Findings
Range("O2") = thiswb.Sheets("TANF Computation").Range("B71") 'Benefit Amt
'Range("P2") = thiswb.Sheets("TANF Computation").Range("C69") 'Benefit Determination
Range("Q2") = thisws.Range("AO10") 'Amt of Error

'reformat street address
    Range("Z2") = StrConv(thisws.Range("B3"), vbProperCase)
'Reformat city and state
    temparry = Split(thisws.Range("B4"), ",")
    TempStr = StrConv(Trim(temparry(LBound(temparry))), vbProperCase)

Range("AA2") = TempStr & ", " & Trim(temparry(UBound(temparry)))
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////
'GA
ElseIf review_type = 9 Then

' Get county number and name in correct format
countynum = Right(thisws.Range("U10"), 2)
districtnum = thisws.Range("Y10")
countylookup = countyformat(countynum, thisws.Range("O4"), districtnum)
Range("AK2") = thisws.Range("M5") & " Assistance Office"
'Look for counties with districts
'If countynum = "02" Or countynum = "23" Or countynum = "40" Or _
'        countynum = "51" Or countynum = "63" Or countynum = "65" Then
'    Range("AK2") = thisws.Range("O4") & " Office"
'Else
'    Range("AK2") = thisws.Range("O4") & " CAO"
'End If

'Client Information
Range("B2") = StrConv(thisws.Range("B2"), vbProperCase) 'Client Name
Range("C2") = Right(thisws.Range("U10"), 2) & "/" & thisws.Range("I10") 'County and Case Number
Range("D2") = thisws.Range("A10") 'Review Number
review_number = Range("D2")
Range("E2") = Left(thisws.Range("AB10"), 2) & "/" & Right(thisws.Range("AB10"), 4) 'Review Month
sample_month = thisws.Range("AB10")
Range("A7") = Val(thisws.Range("AO3") & thisws.Range("AP3")) 'Reviewer Number
Range("G2") = Range("C7") & "/" & thisws.Range("AG2") 'Supervisor and Reviewer Name
Range("E7") = review_type
Range("F7") = review_type
Range("R2") = thisws.Range("J61") & " - " & thisws.Range("O61") & " - " & thisws.Range("T61") 'Element, Nature, & Cause Code
Range("E17") = Val(thisws.Range("AL10")) 'Review Findings
'Range("O2") = thiswb.Sheets("GA Computation").Range("B71") 'Benefit Amt
'Range("P2") = thiswb.Sheets("GA Computation").Range("C71") 'Benefit Determination
Range("Q2") = thisws.Range("AO10") 'Amt of Error

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


End If
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Going out to website with list of Executive Directors
        FullTextFileName = datasourcews.Range("U6")
            
            'ActiveWorkbook.FollowHyperlink (FullTextFileName)
            Workbooks.Open fileName:=FullTextFileName, UpdateLinks:=False, ReadOnly:=True
            
            'Allegheny Districts are formatted as dates, change them to 02/XX format
            For irow = 1 To 100
                If IsDate(Range("A" & irow)) Then
                    TempStr = WorksheetFunction.Text(Range("A" & irow), "mm/d/yyyy")
                    Range("A" & irow).NumberFormat = "@"
                    Range("A" & irow) = Left(TempStr, 4)
                End If
            Next irow
            
            For irow = 1 To 1000
                If Range("A" & irow) = countylookup Then
                    datasourcews.Range("Y2") = Range("D" & irow)
                End If
            Next irow
            ActiveWorkbook.Close False
            
            datasourcewb.Close True
            
   
        findmemo = "AMR for Review Number " & review_number & " for Sample Month " & sample_month & ".doc"
    
    
    'Copy template finding memo to active directory
    SrceFile = PathStr & "Finding Memo\AMR memo.docx"
    tempfindmemoname = sPath & "\FM temp.docx"
    DestFile = tempfindmemoname
    FileCopy SrceFile, DestFile
    
    'Check to see if findings memo has already been created.
    If Dir(sPath & "\" & findmemo) <> "" Then
        Response = MsgBox("An AMR memo has already been created. Do you want to overwrite this memo? No memo will be created if you click NO", Buttons:=vbYesNo)
    ' If statement to check if the NO button was selected.
      If Response = vbNo Then
        End
      End If
    End If

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

    .ActiveDocument.Content.Font.Name = "Arial"
    '.ActiveDocument.Content.Font.Size = 11

    .ActiveDocument.SaveAs sPath & "\" & findmemo
    .ActiveDocument.Close
    
    .Documents(tempfindmemoname).Close SaveChanges:=False

End With
appWD.Quit
Set appWD = Nothing

'delete temp files
'Kill tempfindmemoname
'Kill tempsourcename

Application.StatusBar = False
Application.ScreenUpdating = True
Application.Visible = True

MsgBox "An AMR Memo has been saved as " & sPath & "\" & findmemo & _
". Please close the Internet Explorer with the CAO Directory. - Thank you!"

End Sub
Sub Criminal()

'Create Finding Memo
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

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'TANF
If review_type = 1 Then

' Get county number and name in correct format
countynum = Right(thisws.Range("U10"), 2)
districtnum = thisws.Range("Y10")
countylookup = countyformat(countynum, thisws.Range("O4"), districtnum)
Range("AK2") = thisws.Range("O4") & " Assistance Office"
'Look for counties with districts
'If countynum = "02" Or countynum = "23" Or countynum = "40" Or _
'        countynum = "51" Or countynum = "63" Or countynum = "65" Then
'    Range("AK2") = thisws.Range("O4") & " Office"
'Else
'    Range("AK2") = thisws.Range("O4") & " CAO"
'End If

'Client Information
Range("B2") = StrConv(thisws.Range("B2"), vbProperCase) 'Client Name
Range("C2") = Right(thisws.Range("U10"), 2) & "/" & thisws.Range("I10") 'County and Case Number
Range("D2") = thisws.Range("A10") 'Review Number
review_number = Range("D2")
Range("E2") = Left(thisws.Range("AB10"), 2) & "/" & Right(thisws.Range("AB10"), 4) 'Review Month
sample_month = thisws.Range("AB10")
Range("A7") = Val(thisws.Range("AO3") & thisws.Range("AP3")) 'Reviewer Number
Range("G2") = Range("C7") & "/" & thisws.Range("AG2") 'Supervisor and Reviewer Name
Range("E7") = review_type
Range("F7") = review_type
Range("R2") = thisws.Range("J61") & " - " & thisws.Range("O61") & " - " & thisws.Range("T61") 'Element, Nature, & Cause Code
Range("E17") = Val(thisws.Range("AL10")) 'Review Findings
Range("O2") = thiswb.Sheets("TANF Computation").Range("B72") 'Benefit Amt
Range("P2") = thiswb.Sheets("TANF Computation").Range("C69") 'Benefit Determination
Range("Q2") = thisws.Range("AO10") 'Amt of Error

'reformat street address
    Range("Z2") = StrConv(thisws.Range("B3"), vbProperCase)
'Reformat city and state
    temparry = Split(thisws.Range("B4"), ",")
    TempStr = StrConv(Trim(temparry(LBound(temparry))), vbProperCase)

Range("AA2") = TempStr & ", " & Trim(temparry(UBound(temparry)))
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////
'GA
ElseIf review_type = 9 Then

' Get county number and name in correct format
countynum = Right(thisws.Range("U10"), 2)
districtnum = thisws.Range("Y10")
countylookup = countyformat(countynum, thisws.Range("O4"), districtnum)
Range("AK2") = thisws.Range("M5") & " Assistance Office"
'Look for counties with districts
'If countynum = "02" Or countynum = "23" Or countynum = "40" Or _
'        countynum = "51" Or countynum = "63" Or countynum = "65" Then
'    Range("AK2") = thisws.Range("O4") & " Office"
'Else
'    Range("AK2") = thisws.Range("O4") & " CAO"
'End If

'Client Information
Range("B2") = StrConv(thisws.Range("B2"), vbProperCase) 'Client Name
Range("C2") = Right(thisws.Range("U10"), 2) & "/" & thisws.Range("I10") 'County and Case Number
Range("D2") = thisws.Range("A10") 'Review Number
review_number = Range("D2")
Range("E2") = Left(thisws.Range("AB10"), 2) & "/" & Right(thisws.Range("AB10"), 4) 'Review Month
sample_month = thisws.Range("AB10")
Range("A7") = Val(thisws.Range("AO3") & thisws.Range("AP3")) 'Reviewer Number
Range("G2") = Range("C7") & "/" & thisws.Range("AG2") 'Supervisor and Reviewer Name
Range("E7") = review_type
Range("F7") = review_type
Range("R2") = thisws.Range("J61") & " - " & thisws.Range("O61") & " - " & thisws.Range("T61") 'Element, Nature, & Cause Code
Range("E17") = Val(thisws.Range("AL10")) 'Review Findings
Range("O2") = thiswb.Sheets("GA Computation").Range("B73") 'Benefit Amt
Range("P2") = thiswb.Sheets("GA Computation").Range("C71") 'Benefit Determination
Range("Q2") = thisws.Range("AO10") 'Amt of Error

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


End If
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Going out to website with list of Executive Directors
        FullTextFileName = datasourcews.Range("U6")
            
            'ActiveWorkbook.FollowHyperlink (FullTextFileName)
            Workbooks.Open fileName:=FullTextFileName, UpdateLinks:=False, ReadOnly:=True
            
            'Allegheny Districts are formatted as dates, change them to 02/XX format
            For irow = 1 To 100
                If IsDate(Range("A" & irow)) Then
                    TempStr = WorksheetFunction.Text(Range("A" & irow), "mm/d/yyyy")
                    Range("A" & irow).NumberFormat = "@"
                    Range("A" & irow) = Left(TempStr, 4)
                End If
            Next irow
            
            For irow = 1 To 1000
                If Range("A" & irow) = countylookup Then
                    datasourcews.Range("Y2") = Range("D" & irow)
                End If
            Next irow
            ActiveWorkbook.Close False
            
            datasourcewb.Close True
            
   
        findmemo = "Criminal History Memo for Review Number " & review_number & " for Sample Month " & sample_month & ".doc"
    
    
    'Copy template finding memo to active directory
    SrceFile = PathStr & "Finding Memo\Criminal History.docx"
    tempfindmemoname = sPath & "\FM temp.docx"
    DestFile = tempfindmemoname
    FileCopy SrceFile, DestFile
    
    'Check to see if findings memo has already been created.
    If Dir(sPath & "\" & findmemo) <> "" Then
        Response = MsgBox("A Criminal History memo has already been created. Do you want to overwrite this memo? No memo will be created if you click NO", Buttons:=vbYesNo)
    ' If statement to check if the NO button was selected.
      If Response = vbNo Then
        End
      End If
    End If

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

    .ActiveDocument.Content.Font.Name = "Arial"
    '.ActiveDocument.Content.Font.Size = 11

    .ActiveDocument.SaveAs sPath & "\" & findmemo
    .ActiveDocument.Close
    
    .Documents(tempfindmemoname).Close SaveChanges:=False

End With
appWD.Quit
Set appWD = Nothing

'delete temp files
'Kill tempfindmemoname
'Kill tempsourcename

Application.StatusBar = False
Application.ScreenUpdating = True
Application.Visible = True

MsgBox "A Criminal History Memo has been saved as " & sPath & "\" & findmemo & _
". Please close the Internet Explorer with the CAO Directory. - Thank you!"

End Sub
Sub Info()

'Create Finding Memo
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
SrceFile = PathStr & "Finding Memo/Finding Memo Data Source.xlsx"
DestFile = tempsourcename
FileCopy SrceFile, DestFile

'Open spreadsheet where all the data is stored to populate findings memo
Workbooks.Open fileName:=tempsourcename

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

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'TANF
If review_type = 1 Then

' Get county number and name in correct format
countynum = Right(thisws.Range("U10"), 2)
districtnum = thisws.Range("Y10")
countylookup = countyformat(countynum, thisws.Range("O4"), districtnum)
Range("AK2") = thisws.Range("O4") & " Assistance Office"
'Look for counties with districts
'If countynum = "02" Or countynum = "23" Or countynum = "40" Or _
'        countynum = "51" Or countynum = "63" Or countynum = "65" Then
'    Range("AK2") = thisws.Range("O4") & " Office"
'Else
'    Range("AK2") = thisws.Range("O4") & " CAO"
'End If

'Client Information
Range("B2") = StrConv(thisws.Range("B2"), vbProperCase) 'Client Name
Range("C2") = Right(thisws.Range("U10"), 2) & "/" & thisws.Range("I10") 'County and Case Number
Range("D2") = thisws.Range("A10") 'Review Number
review_number = Range("D2")
Range("E2") = Left(thisws.Range("AB10"), 2) & "/" & Right(thisws.Range("AB10"), 4) 'Review Month
sample_month = thisws.Range("AB10")
Range("A7") = Val(thisws.Range("AO3") & thisws.Range("AP3")) 'Reviewer Number
Range("G2") = Range("C7") & "/" & thisws.Range("AG2") 'Supervisor and Reviewer Name
Range("E7") = review_type
Range("F7") = review_type
Range("R2") = thisws.Range("J61") & " - " & thisws.Range("O61") & " - " & thisws.Range("T61") 'Element, Nature, & Cause Code
Range("E17") = Val(thisws.Range("AL10")) 'Review Findings
Range("O2") = thiswb.Sheets("TANF Computation").Range("B72") 'Benefit Amt
Range("P2") = thiswb.Sheets("TANF Computation").Range("C69") 'Benefit Determination
Range("Q2") = thisws.Range("AO10") 'Amt of Error

'reformat street address
    Range("Z2") = StrConv(thisws.Range("B3"), vbProperCase)
'Reformat city and state
    temparry = Split(thisws.Range("B4"), ",")
    TempStr = StrConv(Trim(temparry(LBound(temparry))), vbProperCase)
    Range("AA2") = TempStr & ", " & Trim(temparry(UBound(temparry)))
    
'category and grant group
Range("CE2") = thisws.Range("Q10")
Range("CG2") = thisws.Range("S10")
'area
areanum = Left(thisws.Range("U10"), 1)
'Column G is TANF Info column in Courtesy Copy for Memos file
cc_col = "G"

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////
'GA
ElseIf review_type = 9 Then

' Get county number and name in correct format
countynum = Right(thisws.Range("U10"), 2)
districtnum = thisws.Range("Y10")
countylookup = countyformat(countynum, thisws.Range("O4"), districtnum)
Range("AK2") = thisws.Range("M5") & " Assistance Office"
'Look for counties with districts
'If countynum = "02" Or countynum = "23" Or countynum = "40" Or _
'        countynum = "51" Or countynum = "63" Or countynum = "65" Then
'    Range("AK2") = thisws.Range("O4") & " Office"
'Else
'    Range("AK2") = thisws.Range("O4") & " CAO"
'End If

'Client Information
Range("B2") = StrConv(thisws.Range("B2"), vbProperCase) 'Client Name
Range("C2") = Right(thisws.Range("U10"), 2) & "/" & thisws.Range("I10") 'County and Case Number
Range("D2") = thisws.Range("A10") 'Review Number
review_number = Range("D2")
Range("E2") = Left(thisws.Range("AB10"), 2) & "/" & Right(thisws.Range("AB10"), 4) 'Review Month
sample_month = thisws.Range("AB10")
Range("A7") = Val(thisws.Range("AO3") & thisws.Range("AP3")) 'Reviewer Number
Range("G2") = Range("C7") & "/" & thisws.Range("AG2") 'Supervisor and Reviewer Name
Range("E7") = review_type
Range("F7") = review_type
Range("R2") = thisws.Range("J61") & " - " & thisws.Range("O61") & " - " & thisws.Range("T61") 'Element, Nature, & Cause Code
Range("E17") = Val(thisws.Range("AL10")) 'Review Findings
Range("O2") = thiswb.Sheets("GA Computation").Range("B73") 'Benefit Amt
Range("P2") = thiswb.Sheets("GA Computation").Range("C71") 'Benefit Determination
Range("Q2") = thisws.Range("AO10") 'Amt of Error

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

'category and grant group
Range("CE2") = thisws.Range("Q10")
Range("CG2") = thisws.Range("S10")
'area
areanum = Left(thisws.Range("U10"), 1)
'Column G is TANF Info column in Courtesy Copy for Memos file
cc_col = "G"

'/////////////////////////////////////////////////////////////////////////////////////////////////////////////

'SNAP Positive
ElseIf review_type = 5 Then

'Client Information
Range("B2") = StrConv(thisws.Range("B4"), vbProperCase)
Range("C2") = thisws.Range("C155") & thisws.Range("D155") & "/" & Left(thisws.Range("I18"), 9)

' Get county number and name in correct format
countynum = thisws.Range("X18")
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
'category
Range("CE2") = "FS"
'area
areanum = Left(thisws.Range("X18"), 1)
'Column E is SNAP Info column in Courtesy Copy for Memos file
cc_col = "E"

'///////////////////////////////////////////////////////////////////////////////////////////////////
'PE
ElseIf Left(thisws.Name, 2) = "24" Then
review_type = 9

' Get county number and name in correct format
countynum = Right(thisws.Range("G15"), 2)
districtnum = thisws.Range("K15")
countylookup = countyformat(countynum, thisws.Range("F2"), districtnum)
Range("AK2") = thisws.Range("F2") & " Assistance Office"

'Client Information
Range("B2") = StrConv(thisws.Range("B7"), vbProperCase)
'reformat street address
If thisws.Range("L7") <> "" Then
    Range("Z2") = StrConv(thisws.Range("L6"), vbProperCase) & ", " & StrConv(thisws.Range("L7"), vbProperCase)
Else
    Range("Z2") = StrConv(thisws.Range("L6"), vbProperCase)
End If

'Reformat city and state
    temparry = Split(thisws.Range("L8"), ",")
    TempStr = StrConv(Trim(temparry(LBound(temparry))), vbProperCase)

Range("AA2") = TempStr & ", " & Trim(temparry(UBound(temparry)))

'Case Number
Range("C2") = Right(thisws.Range("G15"), 2) & "/" & thisws.Range("S15")

'review number
Range("D2") = thisws.Range("L15")
review_number = Range("D2")
'review month
Range("E2") = thisws.Range("T11")
sample_month = thisws.Range("T11")
'reviewer number
Range("A7") = Val(thisws.Range("AB11") & thisws.Range("AC11"))
'supervisor/examiner
Range("G2") = Range("C7") & "/" & Range("BA2")
'Range("R2") = thisws.Range("J61") & " - " & thisws.Range("O61")
'finding
Range("E17") = Val(thisws.Range("AL10"))
Range("F7") = review_type
Range("E7") = review_type
'category and grant group
Range("CE2") = thisws.Range("AB15")
Range("CG2") = thisws.Range("E19")
'area
areanum = Left(thisws.Range("G15"), 1)
'Column C is MA Info column in Courtesy Copy for Memos file
cc_col = "C"

'////////////////////////////////////////////////////////////////////////////////////////////////////////////
'MA Positive
ElseIf review_type = 2 Then

' Get county number and name in correct format
countynum = Right(thisws.Range("AC10"), 2)
districtnum = thisws.Range("AH10")
countylookup = countyformat(countynum, thisws.Range("O4"), districtnum)
Range("AK2") = thisws.Range("O4") & " Assistance Office"
'Look for counties with districts
'If countynum = "02" Or countynum = "23" Or countynum = "40" Or _
'        countynum = "51" Or countynum = "63" Or countynum = "65" Then
'    Range("AK2") = thisws.Range("O4") & " Office"
'Else
'    Range("AK2") = thisws.Range("O4") & " CAO"
'End If

'Client Information
Range("B2") = StrConv(thisws.Range("B2"), vbProperCase) 'Client Name
Range("C2") = Right(thisws.Range("AC10"), 2) & "/" & thisws.Range("H10") 'County and Case Number
Range("D2") = thisws.Range("A10") & thisws.Range("B10") & thisws.Range("C10") & thisws.Range("D10") & thisws.Range("E10") & thisws.Range("F10") 'Review Number
review_number = Range("D2")
Range("E2") = thisws.Range("AM10") 'Review Month
sample_month = thisws.Range("AM10")
Range("A7") = Val(thisws.Range("AO3") & thisws.Range("AP3")) 'Reviewer Number
Range("G2") = Range("C7") & "/" & thisws.Range("AG2") 'Supervisor and Reviewer Name
Range("E7") = review_type
Range("F7") = review_type
'Range("R2") = thisws.Range("J61") & " - " & thisws.Range("O61") & " - " & thisws.Range("T61") 'Element, Nature, & Cause Code
'Range("E17") = Val(thisws.Range("AL10")) 'Review Findings
'Range("O2") = thiswb.Sheets("GA Computation").Range("B73") 'Benefit Amt
'Range("P2") = thiswb.Sheets("GA Computation").Range("C71") 'Benefit Determination
'Range("Q2") = thisws.Range("AO10") 'Amt of Error

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
'category and grant group
Range("CE2") = thisws.Range("Q10")
Range("CG2") = thisws.Range("Z10")
'area
areanum = Left(thisws.Range("AC10"), 1)
'Column C is MA Info column in Courtesy Copy for Memos file
cc_col = "C"

'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'MA Negative
ElseIf review_type = 8 Then

' Get county number and name in correct format
countynum = Right(thisws.Range("G15"), 2)
districtnum = thisws.Range("K15")
countylookup = countyformat(countynum, thisws.Range("F2"), districtnum)
Range("AK2") = thisws.Range("F2") & " Assistance Office"

'Client Information
Range("B2") = StrConv(thisws.Range("B7"), vbProperCase)
'reformat street address
If thisws.Range("L7") <> "" Then
    Range("Z2") = StrConv(thisws.Range("L6"), vbProperCase) & ", " & StrConv(thisws.Range("L7"), vbProperCase)
Else
    Range("Z2") = StrConv(thisws.Range("L6"), vbProperCase)
End If

'Reformat city and state
    temparry = Split(thisws.Range("L8"), ",")
    TempStr = StrConv(Trim(temparry(LBound(temparry))), vbProperCase)

Range("AA2") = TempStr & ", " & Trim(temparry(UBound(temparry)))

'Case Number
Range("C2") = Right(thisws.Range("G15"), 2) & "/" & thisws.Range("S15")

'review number
Range("D2") = thisws.Range("L15")
review_number = Range("D2")
'review month
Range("E2") = thisws.Range("T11")
sample_month = thisws.Range("T11")
'reviewer number
Range("A7") = Val(thisws.Range("AB11") & thisws.Range("AC11"))
'supervisor/examiner
Range("G2") = Range("C7") & "/" & Range("BA2")
'Range("R2") = thisws.Range("J61") & " - " & thisws.Range("O61")
'finding
Range("E17") = Val(thisws.Range("AL10"))
Range("F7") = review_type
Range("E7") = review_type
'category and grant group
Range("CE2") = thisws.Range("AB15")
Range("CG2") = thisws.Range("F19")
'area
areanum = Left(thisws.Range("G15"), 1)
'Column C is MA Info column in Courtesy Copy for Memos file
cc_col = "C"

End If
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Going out to website with list of Executive Directors
        FullTextFileName = datasourcews.Range("U6")
            
            'ActiveWorkbook.FollowHyperlink (FullTextFileName)
            Workbooks.Open fileName:=FullTextFileName, ReadOnly:=True
            
            'Allegheny Districts are formatted as dates, change them to 02/XX format
            For irow = 1 To 100
                If IsDate(Range("A" & irow)) Then
                    TempStr = WorksheetFunction.Text(Range("A" & irow), "mm/d/yyyy")
                    Range("A" & irow).NumberFormat = "@"
                    Range("A" & irow) = Left(TempStr, 4)
                End If
            Next irow
            
            For irow = 1 To 1000
                If Range("A" & irow) = countylookup Then
                    datasourcews.Range("Y2") = Range("D" & irow)
                End If
            Next irow
            ActiveWorkbook.Close False
            
    
'Open CC list file
    Workbooks.Open fileName:=PathStr & "Finding Memo\Courtesy Copy for Memos.xlsx", ReadOnly:=True
    Set ccwb = ActiveWorkbook
    Set ccws = ccwb.Worksheets(areanum)

'Form cc string to add to memo at bottom
    ccstring = vbNewLine & vbNewLine & "cc: " 'first line has cc: then add new line to string
    ccstring = ccstring & ccws.Range(cc_col & "4") & vbNewLine 'Area Manager
    ccstring = ccstring & ccws.Range(cc_col & "5") & vbNewLine 'Area Staff Asst
    'create variable to keep track of which row we are accessing in the Courtesy Copy for Memos file
    'need this variable because Area 1 has an extra line for their QA person
    ccrow = 6
    'check area number, if area 1 need to add QA on cc list
    If areanum = "1" Then
        ccstring = ccstring & ccws.Range(cc_col & ccrow) & vbNewLine 'QA for Area 1
        ccrow = ccrow + 1 'increment row count in Courtesy Copy for Memos file
    End If
    
    'no office managers are listed for info memos
    'ccstring = ccstring & ccws.Range(cc_col & ccrow) & vbNewLine 'Office manager
    ccrow = ccrow + 1 'skip over office manager but need to increment row count
    
   'check for commas in Corrective Action names
    If InStr(ccws.Range(cc_col & ccrow), ",") > 0 Then
        'if comma is found, then split string into parts where commas are
        tempArray = Split(ccws.Range(cc_col & ccrow), ",")
        'loop from lower to upper bound of array to get all of the parts
        For isplit = LBound(tempArray) To UBound(tempArray)
            'put each part of string unto a separate line in CC string
           ccstring = ccstring & Trim(tempArray(isplit)) & vbNewLine 'Corrective Action
        Next isplit
    Else 'no commas were found, add whole line into cc string
        ccstring = ccstring & ccws.Range(cc_col & ccrow) & vbNewLine 'Corrective Action
    End If 'check for commas
    ccrow = ccrow + 1 'increment row count in Courtesy Copy for Memos file
    
    'check for dash or blank in Additional Recipients
    If Not (ccws.Range(cc_col & ccrow) = "-" Or ccws.Range(cc_col & ccrow) = "") Then
        'check for commas in Additional Recipient names
        If InStr(ccws.Range(cc_col & ccrow), ",") > 0 Then
            'if comma is found, then split string into parts where commas are
            tempArray = Split(ccws.Range(cc_col & ccrow), ",")
            'loop from lower to upper bound of array to get all of the parts
            For isplit = LBound(tempArray) To UBound(tempArray)
                'put each part of string unto a separate line in CC string
                ccstring = ccstring & Trim(tempArray(isplit)) & vbNewLine 'Additional Recipients
            Next isplit
        Else 'no commas were found, add whole line into cc string
            ccstring = ccstring & ccws.Range(cc_col & ccrow) & vbNewLine 'Additional Recipients
        End If 'check for commas
    End If 'check for dash or blank
    ccrow = ccrow + 1 'increment row count in Courtesy Copy for Memos file
        
    'check for dash or blank in Program Manager, if dash or blank is found, don't add to ccstring
    If Not (ccws.Range(cc_col & ccrow) = "-" Or ccws.Range(cc_col & ccrow) = "") Then
        ccstring = ccstring & ccws.Range(cc_col & ccrow) & vbNewLine 'Program Manager
    End If
    
    ccrow = ccrow + 1 'increment row count in Courtesy Copy for Memos file
    ccstring = ccstring & ccws.Range(cc_col & ccrow) & vbNewLine  'Posting
    
    ccrow = ccrow + 1 'increment row count in Courtesy Copy for Memos file
    ccstring = ccstring & ccws.Range(cc_col & ccrow) & vbNewLine 'File
    
    'insert cc list into mail merge file
    datasourcews.Range("DF2") = ccstring
    
    'Close cc workbook without saving
    ccwb.Close False
    
    'save and close temp mail merge file
    datasourcewb.Close True
            
    'create name for information memo
    smonth = Replace(sample_month, "/", "_")
    'findmemo = "Information Memo for Review Number " & review_number & " for Sample Month " & smonth & ".docm"
    findmemo = "Information Memo for Review Number " & review_number & " for Sample Month " & smonth & ".docx"
    
    'Copy template info memo to active directory
    SrceFile = PathStr & "Finding Memo\Information Memo.docx"
    tempfindmemoname = sPath & "\FM temp.docx"
    DestFile = tempfindmemoname
    FileCopy SrceFile, DestFile
    
    'Check to see if information memo has already been created.
    If Dir(sPath & "\" & findmemo) <> "" Then
        Response = MsgBox("An Information memo has already been created. Do you want to overwrite this memo? No memo will be created if you click NO", Buttons:=vbYesNo)
    ' If statement to check if the NO button was selected.
      If Response = vbNo Then
        End
      End If
    End If

    'set up Word application
    Set appWD = CreateObject("Word.Application")
    appWD.Visible = True

    'open information template
    With appWD
    .Documents.Open fileName:=tempfindmemoname
    
    Application.StatusBar = "Now creating document"

    'merge data from mail merge file with information memo
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
    
    .ActiveDocument.Content.Font.Name = "Arial"
    '.ActiveDocument.Content.Font.Size = 12
    
    'set name for info memo in same folder as schedule
    tempfilename = sPath & "\" & findmemo
    
    '******************************************************************************************
    '******************************************************************************************
    'This section was to add a button to information memo to copy memo to paperless folder
    'but doesn't seem to work consistently
    '******************************************************************************************
    
    'add a button to the end of the info memo to copy info memo to paperless folder
    'Dim oRng As Word.Range
    '.ActiveDocument.Range.InsertAfter vbCr
    'Set oRng = .ActiveDocument.Range
    'Dim shp As Word.InlineShape
    'oRng.Collapse wdCollapseEnd
    'put button at end
    'Set shp = oRng.InlineShapes.AddOLEControl(ClassType:="Forms.CommandButton.1")
    'make button larger with word wrap
    'shp.OLEFormat.Object.WordWrap = True
    'shp.OLEFormat.Object.FontSize = 10
    'shp.OLEFormat.Object.Height = 56
    'shp.OLEFormat.Object.Width = 56
    'shp.OLEFormat.Object.Caption = "Copy to Paperless"
    
    'Add the procedure to saveas to the paperless folder on the click event of the button
    'change the pathpaperless variable below to save to a different folder
    'Dim sCode As String
    'sCode = "Private Sub " & shp.OLEFormat.Object.Name & "_Click()" & vbCrLf & _
    '        "'Save current document" & vbCrLf & _
    '        "   ActiveDocument.Save" & vbCrLf & _
    '        "'Gets name of document" & vbCrLf & _
    '        "   docname=ActiveDocument.Name" & vbCrLf & _
    '        "'Changes file ending from macro to non-macro document" & vbCrLf & _
    '        "   docname=Replace(docname, "".docm"", "".docx"")" & vbCrLf & _
    '        "'Set path to folder" & vbCrLf & _
    '        "   pathpaperless = ""\\dhs\share\oim\pwimdaubts04\data\stat\dqc\Paperless & Speechless\""" & vbCrLf & _
    '        "   If Dir(pathpaperless, vbDirectory) = """" Then" & vbCrLf & _
    '        "      MsgBox ""The folder Paperless is not on your computer. Stopping program.""" & vbCrLf & _
    '        "      End" & vbCrLf & _
    '        "   End If" & vbCrLf & _
    '        "'Creates new document" & vbCrLf & _
    '        "   Application.Documents.Add ActiveDocument.FullName" & vbCrLf & _
    '        "'Deletes the button on document" & vbCrLf & _
    '        "   For Each o In ActiveDocument.InlineShapes" & vbCrLf & _
    '        "      If o.OLEFormat.Object.Caption = ""Copy to Paperless"" Then" & vbCrLf & _
    '        "          o.Delete" & vbCrLf & _
    '        "      End If" & vbCrLf & _
    '        "   Next" & vbCrLf & _
    '        "'Saves document to desired folder" & vbCrLf & _
    '        "   ActiveDocument.SaveAs2 FileName:=pathpaperless & docname, FileFormat:= wdFormatXMLDocument" & vbCrLf & _
    '        "   ActiveDocument.Close" & vbCrLf & _
    '        "End Sub"
    '.ActiveDocument.VBProject.VBComponents("ThisDocument").CodeModule.AddFromString sCode

    'save information memo with name in macro format
    '.ActiveDocument.SaveAs2 FileName:=tempfilename, _
    '    FileFormat:=wdFormatXMLDocumentMacroEnabled, _
    '    LockComments:=False, _
    '    Password:="", AddToRecentFiles:=True, WritePassword:="", _
    '    ReadOnlyRecommended:=False, EmbedTrueTypeFonts:=False, _
    '    SaveNativePictureFormat:=False, SaveFormsData:=False, SaveAsAOCELetter:= _
    '    False, CompatibilityMode:=14
        
    '    tempmodname = "Module1"
    
    '******************************************************************************************
    '******************************************************************************************
    
    .ActiveDocument.SaveAs fileName:=tempfilename
                
    .ActiveDocument.Close
    
    'close without saving temp document
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

'tell user that informtaion memo has been created
MsgBox "An Information Memo has been saved as " & sPath & "\" & findmemo & "."

End Sub

Sub TANF_Signature_Notification(sig As Integer)

'Create Finding Memo
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
    If LCase(DUNC) = "\\dhs\share\oim\pwimdaubts04\data\stat" Then
        DLetter = "" & oDrives.Item(i) & "\DQC\"
        Exit For
    ElseIf LCase(DUNC) = "\\dhs\share\oim\pwimdaubts04\data\stat\dqc" Then
        DLetter = "" & oDrives.Item(i) & "\"
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
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'TANF

' Get county number and name in correct format
countynum = Right(thisws.Range("U10"), 2)
districtnum = thisws.Range("Y10")
countylookup = countyformat(countynum, thisws.Range("O4"), districtnum)
'Range("AK2") = thisws.Range("O4") & " Assistance Office"
'Look for counties with districts
If countynum = "02" Or countynum = "23" Or countynum = "40" Or _
        countynum = "51" Or countynum = "63" Or countynum = "65" Then
    Range("AK2") = thisws.Range("O4") & " Office"
    director_title = ", District Director"
Else
    Range("AK2") = thisws.Range("O4") & " Assistance Office"
    director_title = ", Executive Director"
End If

'Client Information
Range("B2") = StrConv(thisws.Range("B2"), vbProperCase) 'Client Name
Range("C2") = Right(thisws.Range("U10"), 2) & "/" & thisws.Range("I10") 'County and Case Number
Range("D2") = thisws.Range("A10") 'Review Number
review_number = Range("D2")
Range("E2") = Left(thisws.Range("AB10"), 2) & "/" & Right(thisws.Range("AB10"), 4) 'Review Month
sample_month = thisws.Range("AB10")
Range("A7") = Val(thisws.Range("AO3") & thisws.Range("AP3")) 'Reviewer Number
Range("G2") = Range("C7") & "/" & thisws.Range("AG2") 'Supervisor and Reviewer Name
Range("E7") = review_type
Range("F7") = review_type
Range("R2") = thisws.Range("J61") & " - " & thisws.Range("O61") & " - " & thisws.Range("T61") 'Element, Nature, & Cause Code
Range("E17") = Val(thisws.Range("AL10")) 'Review Findings
Range("O2") = thiswb.Sheets("TANF Computation").Range("B72") 'Benefit Amt
Range("P2") = thiswb.Sheets("TANF Computation").Range("C69") 'Benefit Determination
Range("Q2") = thisws.Range("AO10") 'Amt of Error
Range("CE2") = thisws.Range("Q10") 'category
Range("CG2") = thisws.Range("S10") 'grant group

'reformat street address
    Range("Z2") = StrConv(thisws.Range("B3"), vbProperCase)
'Reformat city and state
    temparry = Split(thisws.Range("B4"), ",")
    TempStr = StrConv(Trim(temparry(LBound(temparry))), vbProperCase)
    Range("AA2") = TempStr & ", " & Trim(temparry(UBound(temparry)))

'if central field office, then use normajean, otherwise use field office manager
If Trim(Range("AC2")) = "Central" Then
    Range("AT2") = "NormaJean Hill, TANF & LIHEAP Program Manager"
Else
    TempStr = Range("AT2")
    Range("AT2") = TempStr & ", Field Office Manager"
End If

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Going out to website with list of Executive Directors
        FullTextFileName = datasourcews.Range("U6")
            
            'ActiveWorkbook.FollowHyperlink (FullTextFileName)
            Workbooks.Open fileName:=FullTextFileName, UpdateLinks:=False, ReadOnly:=True
            
            'Allegheny Districts are formatted as dates, change them to 02/XX format
            For irow = 1 To 100
                If IsDate(Range("A" & irow)) Then
                    TempStr = WorksheetFunction.Text(Range("A" & irow), "mm/d/yyyy")
                    Range("A" & irow).NumberFormat = "@"
                    Range("A" & irow) = Left(TempStr, 4)
                End If
            Next irow
            
            'Look for county or district number
            For l = 1 To 1000
                If Range("A" & l) = countylookup Then
                    datasourcews.Range("Y2") = Range("D" & l) & director_title
                    Exit For
                End If
            Next l
            
                'found_row = Range("A1:A1000").Find(countylookup).Row
                    'Get ED name
                    'datasourcews.Range("Y2") = Range("D" & found_row)
            ActiveWorkbook.Close False
            
            datasourcewb.Close True
            
    'If sig = 1 Then
    '    findmemo = "TANF Signature Notification Info Memo for Review Number " & review_number & " for Sample Month " & sample_month & ".doc"
    '    SrceFile = pathstr & "Finding Memo\Notification Signature Requirement Info Memo.docx"
    'Else
        findmemo = "TANF Notification Info Memo for Review Number " & review_number & " for Sample Month " & sample_month & ".doc"
        SrceFile = PathStr & "Finding Memo\Notification Requirement Info Memo Aug 2013.docx"
    'End If
    
    'Copy template finding memo to active directory
    tempfindmemoname = sPath & "\FM temp.docx"
    DestFile = tempfindmemoname
    FileCopy SrceFile, DestFile
    
    'Check to see if findings memo has already been created.
    If Dir(sPath & "\" & findmemo) <> "" Then
        Response = MsgBox("An info memo has already been created. Do you want to overwrite this memo? No memo will be created if you click NO", Buttons:=vbYesNo)
    ' If statement to check if the NO button was selected.
      If Response = vbNo Then
        End
      End If
    End If

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

    '.ActiveDocument.Content.Font.Name = "Times New Roman"
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

MsgBox "An Info Memo has been saved as " & sPath & "\" & findmemo & _
". Please close the Internet Explorer with the CAO Directory. - Thank you!"

End Sub
Sub MA_Zero()

'Create Finding Memo
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
    If LCase(DUNC) = "\\dhs\share\oim\pwimdaubts04\data\stat" Then
        DLetter = "" & oDrives.Item(i) & "\DQC\"
        Exit For
    ElseIf LCase(DUNC) = "\\dhs\share\oim\pwimdaubts04\data\stat\dqc" Then
        DLetter = "" & oDrives.Item(i) & "\"
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
SrceFile = PathStr & "Finding Memo/Finding Memo Data Source.xlsx"
DestFile = tempsourcename
FileCopy SrceFile, DestFile

'Open spreadsheet where all the data is stored to populate findings memo
Workbooks.Open fileName:=tempsourcename

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

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'MA

'Client Information
Range("B2") = StrConv(thisws.Range("B2"), vbProperCase) 'Client Name
Range("C2") = Right(thisws.Range("AC10"), 2) & "/" & thisws.Range("H10") 'County and Case Number
Range("D2") = thisws.Range("A10") & thisws.Range("B10") & thisws.Range("C10") & thisws.Range("D10") & thisws.Range("E10") & thisws.Range("F10") 'Review Number
review_number = Range("D2")
Range("E2") = thisws.Range("AM10") 'Review Month
sample_month = Left(thisws.Range("AM10"), 2) & Right(thisws.Range("AM10"), 4)
Range("A7") = Val(thisws.Range("AO3") & thisws.Range("AP3")) 'Reviewer Number
Range("G2") = Range("C7") & "/" & thisws.Range("AG2") 'Supervisor and Reviewer Name
Range("E7") = review_type
Range("F7") = review_type

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

 'area
    areanum = Left(thisws.Range("AC10"), 1)
 'Column C is MA Info column in CC for Memos file
 cc_col = "C"
 
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Going out to website with list of Executive Directors
        FullTextFileName = datasourcews.Range("U6")
            
            'ActiveWorkbook.FollowHyperlink (FullTextFileName)
            Workbooks.Open fileName:=FullTextFileName, ReadOnly:=True
            
            'Allegheny Districts are formatted as dates, change them to 02/XX format
            For irow = 1 To 100
                If IsDate(Range("A" & irow)) Then
                    TempStr = WorksheetFunction.Text(Range("A" & irow), "mm/d/yyyy")
                    Range("A" & irow).NumberFormat = "@"
                    Range("A" & irow) = Left(TempStr, 4)
                End If
            Next irow
            
            For irow = 1 To 1000
                If Range("A" & irow) = countylookup Then
                    datasourcews.Range("Y2") = Range("D" & irow)
                End If
            Next irow
            ActiveWorkbook.Close False
    
'Open CC list file
    Workbooks.Open fileName:=PathStr & "Finding Memo\Courtesy Copy for Memos.xlsx", ReadOnly:=True
    Set ccwb = ActiveWorkbook
    Set ccws = ccwb.Worksheets(areanum)

'Form cc string
    ccstring = vbNewLine & vbNewLine & "cc: "
    ccstring = ccstring & ccws.Range(cc_col & "4") & vbNewLine 'Area Manager
    ccstring = ccstring & ccws.Range(cc_col & "5") & vbNewLine 'Area Staff Asst
    ccrow = 6
    'check area number, if area 1 need to add QA on cc list
    If areanum = "1" Then
        ccstring = ccstring & ccws.Range(cc_col & ccrow) & vbNewLine 'QA for Area 1
        ccrow = ccrow + 1
    End If
    
    'no office managers are listed for info memos
    'ccstring = ccstring & ccws.Range(cc_col & ccrow) & vbNewLine 'Office manager
    ccrow = ccrow + 1 'increment row count in Courtesy Copy for Memos file
    
   'check for commas in Corrective Action names
    If InStr(ccws.Range(cc_col & ccrow), ",") > 0 Then
        'if comma is found, then split string into parts where commas are
        tempArray = Split(ccws.Range(cc_col & ccrow), ",")
        'loop from lower to upper bound of array to get all of the parts
        For isplit = LBound(tempArray) To UBound(tempArray)
            'put each part of string unto a separate line in CC string
           ccstring = ccstring & Trim(tempArray(isplit)) & vbNewLine 'Corrective Action
        Next isplit
    Else 'no commas were found, add whole line into cc string
        ccstring = ccstring & ccws.Range(cc_col & ccrow) & vbNewLine 'Corrective Action
    End If 'check for commas
    ccrow = ccrow + 1 'increment row count in Courtesy Copy for Memos file
    
    'check for dash or blank in Additional Recipients
    If Not (ccws.Range(cc_col & ccrow) = "-" Or ccws.Range(cc_col & ccrow) = "") Then
        'check for commas in Additional Recipient names
        If InStr(ccws.Range(cc_col & ccrow), ",") > 0 Then
            'if comma is found, then split string into parts where commas are
            tempArray = Split(ccws.Range(cc_col & ccrow), ",")
            'loop from lower to upper bound of array to get all of the parts
            For isplit = LBound(tempArray) To UBound(tempArray)
                'put each part of string unto a separate line in CC string
                ccstring = ccstring & Trim(tempArray(isplit)) & vbNewLine 'Additional Recipients
            Next isplit
        Else 'no commas were found, add whole line into cc string
            ccstring = ccstring & ccws.Range(cc_col & ccrow) & vbNewLine 'Additional Recipients
        End If 'check for commas
    End If 'check for dash or blank
    ccrow = ccrow + 1 'increment row count in Courtesy Copy for Memos file
        
    'check for dash or blank in Program Manager, if dash or blank is found, don't add to ccstring
    If Not (ccws.Range(cc_col & ccrow) = "-" Or ccws.Range(cc_col & ccrow) = "") Then
        ccstring = ccstring & ccws.Range(cc_col & ccrow) & vbNewLine 'Program Manager
    End If
    
    ccrow = ccrow + 1 'increment row count in Courtesy Copy for Memos file
    ccstring = ccstring & ccws.Range(cc_col & ccrow) & vbNewLine  'Posting
    
    ccrow = ccrow + 1 'increment row count in Courtesy Copy for Memos file
    ccstring = ccstring & ccws.Range(cc_col & ccrow) & vbNewLine 'File
    
    'insert cc list into mail merge file
    datasourcews.Range("DF2") = ccstring
    
    'Close cc workbook without saving
    ccwb.Close False
    
    'save and close temp mail merge file
    datasourcewb.Close True
            
    'create name for information memo
    smonth = Replace(sample_month, "/", "_")
    'findmemo = "Information Memo for Review Number " & review_number & " for Sample Month " & smonth & ".docm"
    findmemo = "Zero Income Request Memo for Review Number " & review_number & " for Sample Month " & smonth & ".docx"
    
    'Copy template info memo to active directory
    SrceFile = PathStr & "Finding Memo\Zero Income Request.docx"
    tempfindmemoname = sPath & "\FM temp.docx"
    DestFile = tempfindmemoname
    FileCopy SrceFile, DestFile
    
    'Check to see if information memo has already been created.
    If Dir(sPath & "\" & findmemo) <> "" Then
        Response = MsgBox("A Zero Income Request memo has already been created. Do you want to overwrite this memo? No memo will be created if you click NO", Buttons:=vbYesNo)
    ' If statement to check if the NO button was selected.
      If Response = vbNo Then
        End
      End If
    End If

    'set up Word application
    Set appWD = CreateObject("Word.Application")
    appWD.Visible = True

    'open information template
    With appWD
    .Documents.Open fileName:=tempfindmemoname
    
    Application.StatusBar = "Now creating document"

    'merge data from mail merge file with information memo
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
    
    .ActiveDocument.Content.Font.Name = "Arial"
    '.ActiveDocument.Content.Font.Size = 12
    
    'set name for info memo in same folder as schedule
    tempfilename = sPath & "\" & findmemo
    
    
    .ActiveDocument.SaveAs fileName:=tempfilename
                
    .ActiveDocument.Close
    
    'close without saving temp document
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

'tell user that informtaion memo has been created
MsgBox "A Zero Income Request Memo has been saved as " & sPath & "\" & findmemo & "."

End Sub
Sub PendLet()

'Create Finding Memo
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
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'MA

' Get county number and name in correct format
countynum = Right(thisws.Range("U10"), 2)
districtnum = thisws.Range("Y10")
countylookup = countyformat(countynum, thisws.Range("O4"), districtnum)

'Look for counties with districts
If countynum = "02" Or countynum = "23" Or countynum = "40" Or _
        countynum = "51" Or countynum = "63" Or countynum = "65" Then
    Range("AK2") = thisws.Range("O4") & " Office"
    director_title = ", District Director"
Else
    Range("AK2") = thisws.Range("O4") & " Assistance Office"
    director_title = ", Executive Director"
End If

'Client Information
Range("B2") = StrConv(thisws.Range("B2"), vbProperCase) 'Client Name
Range("C2") = Right(thisws.Range("AC10"), 2) & "/" & thisws.Range("H10") 'County and Case Number
Range("D2") = thisws.Range("A10") & thisws.Range("B10") & thisws.Range("C10") & thisws.Range("D10") & thisws.Range("E10") & thisws.Range("F10") 'Review Number
review_number = Range("D2")
Range("E2") = thisws.Range("AM10") 'Review Month
sample_month = Left(thisws.Range("AM10"), 2) & Right(thisws.Range("AM10"), 4)
Range("A7") = Val(thisws.Range("AO3") & thisws.Range("AP3")) 'Reviewer Number
Range("G2") = Range("C7") & "/" & thisws.Range("AG2") 'Supervisor and Reviewer Name
Range("E7") = review_type
Range("F7") = review_type


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



'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            datasourcewb.Close True
            
   
        findmemo = "PA472 Pending Letter for Review Number " & review_number & " for Sample Month " & sample_month & ".doc"
        SrceFile = PathStr & "Finding Memo\PA472-Pending Letter.docx"
   
    
    'Copy template finding memo to active directory
    tempfindmemoname = sPath & "\FM temp.docx"
    DestFile = tempfindmemoname
    FileCopy SrceFile, DestFile
    
    'Check to see if findings memo has already been created.
    If Dir(sPath & "\" & findmemo) <> "" Then
        Response = MsgBox("An PA472 Pending Letter has already been created. Do you want to overwrite this memo? No memo will be created if you click NO", Buttons:=vbYesNo)
    ' If statement to check if the NO button was selected.
      If Response = vbNo Then
        End
      End If
    End If

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

    '.ActiveDocument.Content.Font.Name = "Times New Roman"
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

MsgBox "A PA472 Pending Letter has been saved as " & sPath & "\" & findmemo & _
". Please close the Internet Explorer with the CAO Directory. - Thank you!"

End Sub

Sub SelfEmp()

'Create Finding Memo
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
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'MA

' Get county number and name in correct format
countynum = Right(thisws.Range("U10"), 2)
districtnum = thisws.Range("Y10")
countylookup = countyformat(countynum, thisws.Range("O4"), districtnum)

'Look for counties with districts
If countynum = "02" Or countynum = "23" Or countynum = "40" Or _
        countynum = "51" Or countynum = "63" Or countynum = "65" Then
    Range("AK2") = thisws.Range("O4") & " Office"
    director_title = ", District Director"
Else
    Range("AK2") = thisws.Range("O4") & " Assistance Office"
    director_title = ", Executive Director"
End If

'Client Information
Range("B2") = StrConv(thisws.Range("B2"), vbProperCase) 'Client Name
Range("C2") = Right(thisws.Range("AC10"), 2) & "/" & thisws.Range("H10") 'County and Case Number
Range("D2") = thisws.Range("A10") & thisws.Range("B10") & thisws.Range("C10") & thisws.Range("D10") & thisws.Range("E10") & thisws.Range("F10") 'Review Number
review_number = Range("D2")
Range("E2") = thisws.Range("AM10") 'Review Month
sample_month = Left(thisws.Range("AM10"), 2) & Right(thisws.Range("AM10"), 4)
Range("A7") = Val(thisws.Range("AO3") & thisws.Range("AP3")) 'Reviewer Number
Range("G2") = Range("C7") & "/" & thisws.Range("AG2") 'Supervisor and Reviewer Name
Range("E7") = review_type
Range("F7") = review_type


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

thiswb.Activate
lineno = InputBox("Enter Line#:", "Person")
For i = 11 To 20
    If lineno = Sheets("MA Workbook").Cells(i, 10).Value Then
        Name = StrConv(Sheets("MA Workbook").Cells(i, 12).Value, vbProperCase)
    End If
Next i
setype = InputBox("Enter name of self employment:" & vbNewLine & "EXAMPLE: 'babysitting' or 'landscaping'", "Type")
datasourcews.Range("CI2") = Name
datasourcews.Range("CJ2") = setype

'Adds to pending correspondences list (if tracking sheet used)
msg = "Add this letter to Tracking sheet?"
Ans = MsgBox(msg, vbYesNo)

If Ans = vbYes Then
Sheets("MA Tracking").Select
[C20].Select
Do Until ActiveCell.Value = Empty
    ActiveCell.Offset(1, 0).Select
Loop
ActiveCell.Value = "Self Emp Letter for " & "Ln " & lineno
ActiveCell.Offset(0, 3).Value = Date
End If

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            datasourcewb.Close True ' saves the finding memo data source
         
   
        findmemo = "PA472-Self Employment Letter for Review Number " & review_number & " for Sample Month " & sample_month & " for " & Name & ".docx"
        SrceFile = PathStr & "Finding Memo\PA472-SelfEmployment.docx"
   
    
    'Copy template finding memo to active directory
    tempfindmemoname = sPath & "\FM temp.docx"
    DestFile = tempfindmemoname
    FileCopy SrceFile, DestFile
    
    'Check to see if findings memo has already been created.
    If Dir(sPath & "\" & findmemo) <> "" Then
        Response = MsgBox("An PA472 Self Employment Letter has already been created. Do you want to overwrite this memo? No memo will be created if you click NO", Buttons:=vbYesNo)
    ' If statement to check if the NO button was selected.
      If Response = vbNo Then
        End
      End If
    End If

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

    '.ActiveDocument.Content.Font.Name = "Times New Roman"
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

MsgBox "An PA472 Self Employment Letter has been saved as " & sPath & "\" & findmemo & _
". Please close the Internet Explorer with the CAO Directory. - Thank you!"

End Sub
Sub SelfEmpDet()

'Create Finding Memo
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
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'MA

' Get county number and name in correct format
countynum = Right(thisws.Range("U10"), 2)
districtnum = thisws.Range("Y10")
countylookup = countyformat(countynum, thisws.Range("O4"), districtnum)

'Look for counties with districts
If countynum = "02" Or countynum = "23" Or countynum = "40" Or _
        countynum = "51" Or countynum = "63" Or countynum = "65" Then
    Range("AK2") = thisws.Range("O4") & " Office"
    director_title = ", District Director"
Else
    Range("AK2") = thisws.Range("O4") & " Assistance Office"
    director_title = ", Executive Director"
End If

'Client Information
Range("B2") = StrConv(thisws.Range("B2"), vbProperCase) 'Client Name
Range("C2") = Right(thisws.Range("AC10"), 2) & "/" & thisws.Range("H10") 'County and Case Number
Range("D2") = thisws.Range("A10") & thisws.Range("B10") & thisws.Range("C10") & thisws.Range("D10") & thisws.Range("E10") & thisws.Range("F10") 'Review Number
review_number = Range("D2")
Range("E2") = thisws.Range("AM10") 'Review Month
sample_month = Left(thisws.Range("AM10"), 2) & Right(thisws.Range("AM10"), 4)
Range("A7") = Val(thisws.Range("AO3") & thisws.Range("AP3")) 'Reviewer Number
Range("G2") = Range("C7") & "/" & thisws.Range("AG2") 'Supervisor and Reviewer Name
Range("E7") = review_type
Range("F7") = review_type


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
thiswb.Activate
'Adds to pending correspondences list (if tracking sheet used)
msg = "Add this form to Tracking sheet?"
Ans = MsgBox(msg, vbYesNo)

If Ans = vbYes Then
Sheets("MA Tracking").Select
[C20].Select
Do Until ActiveCell.Value = Empty
    ActiveCell.Offset(1, 0).Select
Loop
ActiveCell.Value = "Self Emp Form to client"
ActiveCell.Offset(0, 3).Value = Date
End If


        datasourcewb.Close True ' saves the finding memo data source
         
   
        findmemo = "Self Employment Letter for Review Number " & review_number & " for Sample Month " & sample_month & ".docx"
        SrceFile = PathStr & "Finding Memo\MA Self Employment.docx"
   
    
    'Copy template finding memo to active directory
    tempfindmemoname = sPath & "\FM temp.docx"
    DestFile = tempfindmemoname
    FileCopy SrceFile, DestFile
    
    'Check to see if findings memo has already been created.
    If Dir(sPath & "\" & findmemo) <> "" Then
        Response = MsgBox("A Self Employment Letter has already been created. Do you want to overwrite this memo? No memo will be created if you click NO", Buttons:=vbYesNo)
    ' If statement to check if the NO button was selected.
      If Response = vbNo Then
        End
      End If
    End If

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

    '.ActiveDocument.Content.Font.Name = "Times New Roman"
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

MsgBox "A Self Employment Letter has been saved as " & sPath & "\" & findmemo & _
". Please close the Internet Explorer with the CAO Directory. - Thank you!"

End Sub

Sub PA76()

'Create Finding Memo
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
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'MA

' Get county number and name in correct format
countynum = Right(thisws.Range("U10"), 2)
districtnum = thisws.Range("Y10")
countylookup = countyformat(countynum, thisws.Range("O4"), districtnum)

'Look for counties with districts
If countynum = "02" Or countynum = "23" Or countynum = "40" Or _
        countynum = "51" Or countynum = "63" Or countynum = "65" Then
    Range("AK2") = thisws.Range("O4") & " Office"
    director_title = ", District Director"
Else
    Range("AK2") = thisws.Range("O4") & " Assistance Office"
    director_title = ", Executive Director"
End If

'Client Information
Range("B2") = StrConv(thisws.Range("B2"), vbProperCase) 'Client Name
Range("C2") = Right(thisws.Range("AC10"), 2) & "/" & thisws.Range("H10") 'County and Case Number
Range("D2") = thisws.Range("A10") & thisws.Range("B10") & thisws.Range("C10") & thisws.Range("D10") & thisws.Range("E10") & thisws.Range("F10") 'Review Number
review_number = Range("D2")
Range("E2") = thisws.Range("AM10") 'Review Month
sample_month = Left(thisws.Range("AM10"), 2) & Right(thisws.Range("AM10"), 4)
Range("A7") = Val(thisws.Range("AO3") & thisws.Range("AP3")) 'Reviewer Number
Range("G2") = Range("C7") & "/" & thisws.Range("AG2") 'Supervisor and Reviewer Name
Range("E7") = review_type
Range("F7") = review_type


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
'Activates MA Schedule to look at workbook
thiswb.Activate
'will ask reviewer what line number they are referring to for the memo
lineno = InputBox("Enter Line#:", "Person")
For i = 11 To 20
    If lineno = Sheets("MA Workbook").Cells(i, 10).Value Then
        Name = StrConv(Sheets("MA Workbook").Cells(i, 12).Value, vbProperCase)
        firstname = StrConv(Trim(Left(Name, InStr(Name, " ") - 1)), vbProperCase)
        ssn = "XXX-XX-" & Right(Sheets("MA Workbook").Cells(i, 31), 4)
        birth = Sheets("MA Workbook").Cells(i, 22).Value
    End If
Next i

bankname = InputBox("Enter name of bank:", "bank")
bankaddress1 = InputBox("Enter 1st line of bank address:", "Address")
bankaddress2 = InputBox("Enter 2nd line of bank address:", "Address")
bankaddress3 = InputBox("Enter 3rd line of bank address:", "Address")


datasourcews.Range("CI2") = Name
datasourcews.Range("CQ2") = ssn
datasourcews.Range("CR2") = birth
datasourcews.Range("CS2") = firstname
datasourcews.Range("CX2") = bankname
datasourcews.Range("CY2") = bankaddress1
datasourcews.Range("CZ2") = bankaddress2
datasourcews.Range("DA2") = bankaddress3

'Adds to pending correspondences list (if tracking sheet used)
msg = "Add this bank to Tracking sheet?"
Ans = MsgBox(msg, vbYesNo)

If Ans = vbYes Then
Sheets("MA Tracking").Select
[C20].Select
Do Until ActiveCell.Value = Empty
    ActiveCell.Offset(1, 0).Select
Loop
ActiveCell.Value = "PA76 to " & bankname & " for Ln " & lineno
ActiveCell.Offset(0, 3).Value = Date
End If

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            datasourcewb.Close True ' saves the finding memo data source
         
   
        findmemo = "PA76 Form for Review Number " & review_number & " for Sample Month " & sample_month & " " & Name & "_" & bankname & ".docx"
        SrceFile = PathStr & "Finding Memo\PA76.docx"
   
    
    'Copy template finding memo to active directory
    tempfindmemoname = sPath & "\FM temp.docx"
    DestFile = tempfindmemoname
    FileCopy SrceFile, DestFile
    
    'Check to see if findings memo has already been created.
    If Dir(sPath & "\" & findmemo) <> "" Then
        Response = MsgBox("A PA76 form has already been created. Do you want to overwrite this memo? No memo will be created if you click NO", Buttons:=vbYesNo)
    ' If statement to check if the NO button was selected.
      If Response = vbNo Then
        End
      End If
    End If

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

MsgBox "A PA76 Form has been saved as " & sPath & "\" & findmemo & _
". Please close the Internet Explorer with the CAO Directory. - Thank you!"

End Sub

Sub PA78()

'Create Finding Memo
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
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'MA

' Get county number and name in correct format
countynum = Right(thisws.Range("U10"), 2)
districtnum = thisws.Range("Y10")
countylookup = countyformat(countynum, thisws.Range("O4"), districtnum)

'Look for counties with districts
If countynum = "02" Or countynum = "23" Or countynum = "40" Or _
        countynum = "51" Or countynum = "63" Or countynum = "65" Then
    Range("AK2") = thisws.Range("O4") & " Office"
    director_title = ", District Director"
Else
    Range("AK2") = thisws.Range("O4") & " Assistance Office"
    director_title = ", Executive Director"
End If

'Client Information
Range("B2") = StrConv(thisws.Range("B2"), vbProperCase) 'Client Name
Range("C2") = Right(thisws.Range("AC10"), 2) & "/" & thisws.Range("H10") 'County and Case Number
Range("D2") = thisws.Range("A10") & thisws.Range("B10") & thisws.Range("C10") & thisws.Range("D10") & thisws.Range("E10") & thisws.Range("F10") 'Review Number
review_number = Range("D2")
Range("E2") = thisws.Range("AM10") 'Review Month
sample_month = Left(thisws.Range("AM10"), 2) & Right(thisws.Range("AM10"), 4)
Range("A7") = Val(thisws.Range("AO3") & thisws.Range("AP3")) 'Reviewer Number
Range("G2") = Range("C7") & "/" & thisws.Range("AG2") 'Supervisor and Reviewer Name
Range("E7") = review_type
Range("F7") = review_type
Range("CE2") = thisws.Range("Q10") & "-" & thisws.Range("V10")


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
'Activates MA Schedule to look at workbook
thiswb.Activate
'will ask reviewer what line number they are referring to for the memo
lineno = InputBox("Enter Line#:", "Person")
For i = 11 To 20
    If lineno = Sheets("MA Workbook").Cells(i, 10).Value Then
        Name = StrConv(Sheets("MA Workbook").Cells(i, 12).Value, vbProperCase)
        ssn = "XXX-XX-" & Right(Sheets("MA Workbook").Cells(i, 31), 4)
        birth = Sheets("MA Workbook").Cells(i, 22).Value
    End If
Next i



employername = InputBox("Enter Name of Employer:", "Employer")
employeraddress1 = InputBox("Enter 1st line of address:", "Address")
employeraddress2 = InputBox("Enter 2nd line of address:", "Address")
employeraddress3 = InputBox("Enter 3rd line of address:", "Address")


datasourcews.Range("CI2") = Name
datasourcews.Range("CQ2") = ssn
datasourcews.Range("CR2") = birth
datasourcews.Range("CT2") = employername
datasourcews.Range("CU2") = employeraddress1
datasourcews.Range("CV2") = employeraddress2
datasourcews.Range("CW2") = employeraddress3

'Adds to pending correspondences list (if tracking sheet used)
msg = "Add this employer to Tracking sheet?"
Ans = MsgBox(msg, vbYesNo)

If Ans = vbYes Then
Sheets("MA Tracking").Select
[C20].Select
Do Until ActiveCell.Value = Empty
    ActiveCell.Offset(1, 0).Select
Loop
ActiveCell.Value = "PA78 to " & employername & " for Ln " & lineno
ActiveCell.Offset(0, 3).Value = Date
End If

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            datasourcewb.Close True ' saves the finding memo data source
         
   
        findmemo = "PA78 Form for Review Number " & review_number & " for Sample Month " & sample_month & " " & Name & "_" & employername & ".docx"
        SrceFile = PathStr & "Finding Memo\PA78.docx"
   
    
    'Copy template finding memo to active directory
    tempfindmemoname = sPath & "\FM temp.docx"
    DestFile = tempfindmemoname
    FileCopy SrceFile, DestFile
    
    'Check to see if findings memo has already been created.
    If Dir(sPath & "\" & findmemo) <> "" Then
        Response = MsgBox("A PA78 form has already been created. Do you want to overwrite this memo? No memo will be created if you click NO", Buttons:=vbYesNo)
    ' If statement to check if the NO button was selected.
      If Response = vbNo Then
        End
      End If
    End If

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

MsgBox "A PA78 Form has been saved as " & sPath & "\" & findmemo & _
". Please close the Internet Explorer with the CAO Directory. - Thank you!"

End Sub
Sub SNAPPA78()

'Create Finding Memo
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
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'SNAP

' Get county number and name in correct format
countynum = Right(thisws.Range("X18"), 2)
districtnum = thisws.Range("B153") & thisws.Range("C153")
countylookup = countyformat(countynum, thisws.Range("M5"), districtnum)
Range("AK2") = thisws.Range("M5") & " Assistance Office"


'Client Information
Range("B2") = StrConv(thisws.Range("B4"), vbProperCase) 'Client Name
Range("C2") = Right(thisws.Range("X18"), 2) & "/" & thisws.Range("I18") 'County and Case Number
Range("D2") = thisws.Range("A18") 'Review Number
review_number = Range("D2")
Range("E2") = thisws.Range("AD18") & "/" & thisws.Range("AG18") 'Review Month
sample_month = thisws.Range("AD18") & thisws.Range("AG18")
Range("A7") = Val(thisws.Range("AJ5") & thisws.Range("AK5")) 'Reviewer Number
Range("G2") = Range("C7") & "/" & thisws.Range("AE5") 'Supervisor and Reviewer Name
Range("E7") = review_type
Range("F7") = review_type

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
'Activates MA Schedule to look at workbook
thiswb.Activate
'will ask reviewer what line number they are referring to for the memo
lineno = InputBox("Enter Line#:", "Person")
For i = 11 To 22
    If lineno = Sheets("FS Workbook").Cells(i, 11).Value Then
        Name = StrConv(Sheets("FS Workbook").Cells(i, 13).Value, vbProperCase)
        ssn = "XXX-XX-" & Right(Sheets("FS Workbook").Cells(i, 32), 4)
        birth = Sheets("FS Workbook").Cells(i, 25).Value
    End If
Next i


employername = InputBox("Enter Name of Employer:", "Employer")
employeraddress1 = InputBox("Enter 1st line of address:", "Address")
employeraddress2 = InputBox("Enter 2nd line of address:", "Address")
employeraddress3 = InputBox("Enter 3rd line of address:", "Address")


datasourcews.Range("CI2") = Name
datasourcews.Range("CQ2") = ssn
datasourcews.Range("CR2") = birth
datasourcews.Range("CT2") = employername
datasourcews.Range("CU2") = employeraddress1
datasourcews.Range("CV2") = employeraddress2
datasourcews.Range("CW2") = employeraddress3


'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            datasourcewb.Close True ' saves the finding memo data source
         
   
        findmemo = "PA78 Form for Review Number " & review_number & " for Sample Month " & sample_month & " " & Name & "_" & employername & ".docx"
        SrceFile = PathStr & "Finding Memo\PA78.docx"
   
    
    'Copy template finding memo to active directory
    tempfindmemoname = sPath & "\FM temp.docx"
    DestFile = tempfindmemoname
    FileCopy SrceFile, DestFile
    
    'Check to see if findings memo has already been created.
    If Dir(sPath & "\" & findmemo) <> "" Then
        Response = MsgBox("A PA78 form has already been created. Do you want to overwrite this memo? No memo will be created if you click NO", Buttons:=vbYesNo)
    ' If statement to check if the NO button was selected.
      If Response = vbNo Then
        End
      End If
    End If

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

MsgBox "A PA78 Form has been saved as " & sPath & "\" & findmemo & _
". Please close the Internet Explorer with the CAO Directory. - Thank you!"

End Sub

Sub TANFPA78()

'Create Finding Memo
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
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'TANF

' Get county number and name in correct format
countynum = Right(thisws.Range("U10"), 2)
districtnum = thisws.Range("Y10")
countylookup = countyformat(countynum, thisws.Range("O4"), districtnum)

'Look for counties with districts
If countynum = "02" Or countynum = "23" Or countynum = "40" Or _
        countynum = "51" Or countynum = "63" Or countynum = "65" Then
    Range("AK2") = thisws.Range("O4") & " Office"
    director_title = ", District Director"
Else
    Range("AK2") = thisws.Range("O4") & " Assistance Office"
    director_title = ", Executive Director"
End If

'Client Information
Range("B2") = StrConv(thisws.Range("B2"), vbProperCase) 'Client Name
Range("C2") = Right(thisws.Range("U10"), 2) & "/" & thisws.Range("I10") 'County and Case Number
Range("D2") = thisws.Range("A10") 'Review Number
review_number = Range("D2")
Range("E2") = Left(thisws.Range("AB10"), 2) & "/" & Right(thisws.Range("AB10"), 4) 'Review Month
sample_month = thisws.Range("AB10")
Range("A7") = Val(thisws.Range("AO3") & thisws.Range("AP3")) 'Reviewer Number
Range("G2") = Range("C7") & "/" & thisws.Range("AG2") 'Supervisor and Reviewer Name
Range("E7") = review_type
Range("F7") = review_type
Range("CE2") = thisws.Range("Q10") 'category


'reformat street address
'If thisws.Range("B4") <> "" Then
'    Range("Z2") = StrConv(thisws.Range("B3"), vbProperCase) & ", " & StrConv(thisws.Range("B4"), vbProperCase)
'Else
    Range("Z2") = StrConv(thisws.Range("B3"), vbProperCase)
'End If


'Reformat city and state
If review_type = 1 Then
    temparry = Split(thisws.Range("B4"), ",")
    TempStr = StrConv(Trim(temparry(LBound(temparry))), vbProperCase)
Else
    temparry = Split(thisws.Range("B5"), ",")
    TempStr = StrConv(Trim(temparry(LBound(temparry))), vbProperCase)
End If

Range("AA2") = TempStr & ", " & Trim(temparry(UBound(temparry)))
'Activates MA Schedule to look at workbook
thiswb.Activate
'will ask reviewer what line number they are referring to for the memo
If review_type = 1 Then
lineno = InputBox("Enter Line#:", "Person")
For i = 11 To 20
    If lineno = Sheets("TANF Workbook").Cells(i, 10).Value Then
        Name = StrConv(Sheets("TANF Workbook").Cells(i, 12).Value, vbProperCase)
        ssn = "XXX-XX-" & Right(Sheets("TANF Workbook").Cells(i, 31), 4)
        birth = Sheets("TANF Workbook").Cells(i, 22).Value
    End If
Next i
Else
lineno = InputBox("Enter Line#:", "Person")
For i = 11 To 20
    If lineno = Sheets("GA Workbook").Cells(i, 10).Value Then
        Name = StrConv(Sheets("GA Workbook").Cells(i, 12).Value, vbProperCase)
        ssn = "XXX-XX-" & Right(Sheets("GA Workbook").Cells(i, 31), 4)
        birth = Sheets("GA Workbook").Cells(i, 22).Value
    End If
Next i
End If



employername = InputBox("Enter Name of Employer:", "Employer")
employeraddress1 = InputBox("Enter 1st line of address:", "Address")
employeraddress2 = InputBox("Enter 2nd line of address:", "Address")
employeraddress3 = InputBox("Enter 3rd line of address:", "Address")


datasourcews.Range("CI2") = Name
datasourcews.Range("CQ2") = ssn
datasourcews.Range("CR2") = birth
datasourcews.Range("CT2") = employername
datasourcews.Range("CU2") = employeraddress1
datasourcews.Range("CV2") = employeraddress2
datasourcews.Range("CW2") = employeraddress3

'Adds to pending correspondences list (if tracking sheet used)
'Msg = "Add this employer to Tracking sheet?"
'Ans = MsgBox(Msg, vbYesNo)

'If Ans = vbYes Then
'Sheets("MA Tracking").Select
'[C20].Select
'Do Until ActiveCell.value = Empty
'    ActiveCell.Offset(1, 0).Select
'Loop
'ActiveCell.value = "PA78 to " & employername & " for Ln " & lineno
'ActiveCell.Offset(0, 3).value = Date
'End If

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            datasourcewb.Close True ' saves the finding memo data source
         
   
        findmemo = "Employment and Earnings Report Form for Review Number " & review_number & " for Sample Month " & sample_month & " " & Name & "_" & employername & ".docx"
        SrceFile = PathStr & "Finding Memo\TANFPA78.docx"
   
    
    'Copy template finding memo to active directory
    tempfindmemoname = sPath & "\FM temp.docx"
    DestFile = tempfindmemoname
    FileCopy SrceFile, DestFile
    
    'Check to see if findings memo has already been created.
    If Dir(sPath & "\" & findmemo) <> "" Then
        Response = MsgBox("An Employment/Earnings Report form has already been created. Do you want to overwrite this memo? No memo will be created if you click NO", Buttons:=vbYesNo)
    ' If statement to check if the NO button was selected.
      If Response = vbNo Then
        End
      End If
    End If

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

MsgBox "An Employment/Earnings Report has been saved as " & sPath & "\" & findmemo & _
". Please close the Internet Explorer with the CAO Directory. - Thank you!"

End Sub
Sub MA_Supp()

'Create Finding Memo
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
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'MA

' Get county number and name in correct format
countynum = Right(thisws.Range("U10"), 2)
districtnum = thisws.Range("Y10")
countylookup = countyformat(countynum, thisws.Range("O4"), districtnum)

'Look for counties with districts
If countynum = "02" Or countynum = "23" Or countynum = "40" Or _
        countynum = "51" Or countynum = "63" Or countynum = "65" Then
    Range("AK2") = thisws.Range("O4") & " Office"
    director_title = ", District Director"
Else
    Range("AK2") = thisws.Range("O4") & " Assistance Office"
    director_title = ", Executive Director"
End If

'Client Information
Range("B2") = StrConv(thisws.Range("B2"), vbProperCase) 'Client Name
Range("C2") = Right(thisws.Range("AC10"), 2) & "/" & thisws.Range("H10") 'County and Case Number
Range("D2") = thisws.Range("A10") & thisws.Range("B10") & thisws.Range("C10") & thisws.Range("D10") & thisws.Range("E10") & thisws.Range("F10") 'Review Number
review_number = Range("D2")
Range("E2") = thisws.Range("AM10") 'Review Month
sample_month = Left(thisws.Range("AM10"), 2) & Right(thisws.Range("AM10"), 4)
Range("A7") = Val(thisws.Range("AO3") & thisws.Range("AP3")) 'Reviewer Number
Range("G2") = Range("C7") & "/" & thisws.Range("AG2") 'Supervisor and Reviewer Name
Range("E7") = review_type
Range("F7") = review_type

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



'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            datasourcewb.Close True ' saves the finding memo data source
         
   
        findmemo = "MA Support Form for Review Number " & review_number & " for Sample Month " & sample_month & " " & Name & "_" & bankname & ".docx"
        SrceFile = PathStr & "Finding Memo\MA Support Form.docx"
   
    
    'Copy template finding memo to active directory
    tempfindmemoname = sPath & "\FM temp.docx"
    DestFile = tempfindmemoname
    FileCopy SrceFile, DestFile
    
    'Check to see if findings memo has already been created.
    If Dir(sPath & "\" & findmemo) <> "" Then
        Response = MsgBox("A MA Support Form has already been created. Do you want to overwrite this memo? No memo will be created if you click NO", Buttons:=vbYesNo)
    ' If statement to check if the NO button was selected.
      If Response = vbNo Then
        End
      End If
    End If

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

MsgBox "A MA Support Form has been saved as " & sPath & "\" & findmemo & _
". Please close the Internet Explorer with the CAO Directory. - Thank you!"

End Sub
Sub PA83Z()

'Create Finding Memo
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
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'MA

' Get county number and name in correct format
countynum = Right(thisws.Range("U10"), 2)
districtnum = thisws.Range("Y10")
countylookup = countyformat(countynum, thisws.Range("O4"), districtnum)

'Look for counties with districts
If countynum = "02" Or countynum = "23" Or countynum = "40" Or _
        countynum = "51" Or countynum = "63" Or countynum = "65" Then
    Range("AK2") = thisws.Range("O4") & " Office"
    director_title = ", District Director"
Else
    Range("AK2") = thisws.Range("O4") & " Assistance Office"
    director_title = ", Executive Director"
End If

'Client Information
Range("B2") = StrConv(thisws.Range("B2"), vbProperCase) 'Client Name
Range("C2") = Right(thisws.Range("AC10"), 2) & "/" & thisws.Range("H10") 'County and Case Number
Range("D2") = thisws.Range("A10") & thisws.Range("B10") & thisws.Range("C10") & thisws.Range("D10") & thisws.Range("E10") & thisws.Range("F10") 'Review Number
review_number = Range("D2")
Range("E2") = thisws.Range("AM10") 'Review Month
sample_month = Left(thisws.Range("AM10"), 2) & Right(thisws.Range("AM10"), 4)
Range("A7") = Val(thisws.Range("AO3") & thisws.Range("AP3")) 'Reviewer Number
Range("G2") = Range("C7") & "/" & thisws.Range("AG2") 'Supervisor and Reviewer Name
Range("E7") = review_type
Range("F7") = review_type


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
'Activates MA Schedule to look at workbook
thiswb.Activate
'will ask reviewer what line number they are referring to for the memo
lineno = InputBox("Enter Line#:", "Person")
For i = 11 To 20
    If lineno = Sheets("MA Workbook").Cells(i, 10).Value Then
        Name = StrConv(Sheets("MA Workbook").Cells(i, 12).Value, vbProperCase)
        firstname = StrConv(Trim(Left(Name, InStr(Name, " ") - 1)), vbProperCase)
        ssn = "XXX-XX-" & Right(Sheets("MA Workbook").Cells(i, 31), 4)
    End If
Next i

bankname = InputBox("Enter name of bank:", "bank")
bankaddress1 = InputBox("Enter 1st line of bank address:", "Address")
bankaddress2 = InputBox("Enter 2nd line of bank address:", "Address")
bankaddress3 = InputBox("Enter 3rd line of bank address:", "Address")

policy = InputBox("Enter first known policy number:", "Policy")

datasourcews.Range("CI2") = Name
datasourcews.Range("CQ2") = ssn
datasourcews.Range("CS2") = firstname
datasourcews.Range("CX2") = bankname
datasourcews.Range("CY2") = bankaddress1
datasourcews.Range("CZ2") = bankaddress2
datasourcews.Range("DA2") = bankaddress3
datasourcews.Range("DB2") = policy

'Adds to pending correspondences list (if tracking sheet used)
msg = "Add this financial institution to Tracking sheet?"
Ans = MsgBox(msg, vbYesNo)

If Ans = vbYes Then
Sheets("MA Tracking").Select
[C20].Select
Do Until ActiveCell.Value = Empty
    ActiveCell.Offset(1, 0).Select
Loop
ActiveCell.Value = "PA83-Z to " & bankname & " for Ln " & lineno
ActiveCell.Offset(0, 3).Value = Date
End If

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            datasourcewb.Close True ' saves the finding memo data source
         
   
        findmemo = "PA83-Z Form for Review Number " & review_number & " for Sample Month " & sample_month & " " & Name & "_" & bankname & ".docx"
        SrceFile = PathStr & "Finding Memo\PA83Z.docx"
   
    
    'Copy template finding memo to active directory
    tempfindmemoname = sPath & "\FM temp.docx"
    DestFile = tempfindmemoname
    FileCopy SrceFile, DestFile
    
    'Check to see if findings memo has already been created.
    If Dir(sPath & "\" & findmemo) <> "" Then
        Response = MsgBox("A PA83Z form has already been created. Do you want to overwrite this memo? No memo will be created if you click NO", Buttons:=vbYesNo)
    ' If statement to check if the NO button was selected.
      If Response = vbNo Then
        End
      End If
    End If

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

MsgBox "A PA83Z Form has been saved as " & sPath & "\" & findmemo & _
". Please close the Internet Explorer with the CAO Directory. - Thank you!"

End Sub

Sub HouseholdComp()

'Create Finding Memo
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
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'MA

'Client Information
Range("B2") = StrConv(thisws.Range("B2"), vbProperCase) 'Client Name
Range("C2") = Right(thisws.Range("AC10"), 2) & "/" & thisws.Range("H10") 'County and Case Number
Range("D2") = thisws.Range("A10") & thisws.Range("B10") & thisws.Range("C10") & thisws.Range("D10") & thisws.Range("E10") & thisws.Range("F10") 'Review Number
review_number = Range("D2")
Range("E2") = thisws.Range("AM10") 'Review Month
sample_month = Left(thisws.Range("AM10"), 2) & Right(thisws.Range("AM10"), 4)
Range("A7") = Val(thisws.Range("AO3") & thisws.Range("AP3")) 'Reviewer Number
Range("G2") = Range("C7") & "/" & thisws.Range("AG2") 'Supervisor and Reviewer Name
Range("E7") = review_type
Range("F7") = review_type


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


'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            datasourcewb.Close True ' saves the finding memo data source
         
   
        findmemo = "Household Composition Form for Review Number " & review_number & " for Sample Month " & sample_month & ".docx"
        SrceFile = PathStr & "Finding Memo\Household Composition.docx"
   
    
    'Copy template finding memo to active directory
    tempfindmemoname = sPath & "\FM temp.docx"
    DestFile = tempfindmemoname
    FileCopy SrceFile, DestFile
    
    'Check to see if findings memo has already been created.
    If Dir(sPath & "\" & findmemo) <> "" Then
        Response = MsgBox("A Household Composition form has already been created. Do you want to overwrite this memo? No memo will be created if you click NO", Buttons:=vbYesNo)
    ' If statement to check if the NO button was selected.
      If Response = vbNo Then
        End
      End If
    End If

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

MsgBox "A Household Composition Form has been saved as " & sPath & "\" & findmemo & "."


End Sub
Sub Threshold()

'Create Finding Memo
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
SrceFile = PathStr & "Finding Memo/Finding Memo Data Source.xlsx"
DestFile = tempsourcename
FileCopy SrceFile, DestFile

'Open spreadsheet where all the data is stored to populate findings memo
Workbooks.Open fileName:=tempsourcename

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

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////

'SNAP Positive


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
'category
Range("CE2") = "FS"
'area
areanum = Left(thisws.Range("X18"), 1)
'Column E is SNAP Info column in Courtesy Copy for Memos file
cc_col = "E"


'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Going out to website with list of Executive Directors
        FullTextFileName = datasourcews.Range("U6")
            
            'ActiveWorkbook.FollowHyperlink (FullTextFileName)
            Workbooks.Open fileName:=FullTextFileName, ReadOnly:=True
            
            'Allegheny Districts are formatted as dates, change them to 02/XX format
            For irow = 1 To 100
                If IsDate(Range("A" & irow)) Then
                    TempStr = WorksheetFunction.Text(Range("A" & irow), "mm/d/yyyy")
                    Range("A" & irow).NumberFormat = "@"
                    Range("A" & irow) = Left(TempStr, 4)
                End If
            Next irow
            
            For irow = 1 To 1000
                If Range("A" & irow) = countylookup Then
                    datasourcews.Range("Y2") = Range("D" & irow)
                End If
            Next irow
            ActiveWorkbook.Close False
            
    
'Open CC list file
    Workbooks.Open fileName:=PathStr & "Finding Memo\Courtesy Copy for Memos.xlsx", ReadOnly:=True
    Set ccwb = ActiveWorkbook
    Set ccws = ccwb.Worksheets(areanum)

'Form cc string
    ccstring = vbNewLine & vbNewLine & "cc: "
    ccstring = ccstring & ccws.Range(cc_col & "4") & vbNewLine 'Area Manager
    ccstring = ccstring & ccws.Range(cc_col & "5") & vbNewLine 'Area Staff Asst
    ccrow = 6
    'check area number, if area 1 need to add QA on cc list
    If areanum = "1" Then
        ccstring = ccstring & ccws.Range(cc_col & ccrow) & vbNewLine 'QA for Area 1
        ccrow = ccrow + 1
    End If
    
    'no office managers are listed for info memos
    'ccstring = ccstring & ccws.Range(cc_col & ccrow) & vbNewLine 'Office manager
    ccrow = ccrow + 1 'increment row count in Courtesy Copy for Memos file
    
   'check for commas in Corrective Action names
    If InStr(ccws.Range(cc_col & ccrow), ",") > 0 Then
        'if comma is found, then split string into parts where commas are
        tempArray = Split(ccws.Range(cc_col & ccrow), ",")
        'loop from lower to upper bound of array to get all of the parts
        For isplit = LBound(tempArray) To UBound(tempArray)
            'put each part of string unto a separate line in CC string
           ccstring = ccstring & Trim(tempArray(isplit)) & vbNewLine 'Corrective Action
        Next isplit
    Else 'no commas were found, add whole line into cc string
        ccstring = ccstring & ccws.Range(cc_col & ccrow) & vbNewLine 'Corrective Action
    End If 'check for commas
    ccrow = ccrow + 1 'increment row count in Courtesy Copy for Memos file
    
    'check for dash or blank in Additional Recipients
    If Not (ccws.Range(cc_col & ccrow) = "-" Or ccws.Range(cc_col & ccrow) = "") Then
        'check for commas in Additional Recipient names
        If InStr(ccws.Range(cc_col & ccrow), ",") > 0 Then
            'if comma is found, then split string into parts where commas are
            tempArray = Split(ccws.Range(cc_col & ccrow), ",")
            'loop from lower to upper bound of array to get all of the parts
            For isplit = LBound(tempArray) To UBound(tempArray)
                'put each part of string unto a separate line in CC string
                ccstring = ccstring & Trim(tempArray(isplit)) & vbNewLine 'Additional Recipients
            Next isplit
        Else 'no commas were found, add whole line into cc string
            ccstring = ccstring & ccws.Range(cc_col & ccrow) & vbNewLine 'Additional Recipients
        End If 'check for commas
    End If 'check for dash or blank
    ccrow = ccrow + 1 'increment row count in Courtesy Copy for Memos file
        
    'check for dash or blank in Program Manager, if dash or blank is found, don't add to ccstring
    If Not (ccws.Range(cc_col & ccrow) = "-" Or ccws.Range(cc_col & ccrow) = "") Then
        ccstring = ccstring & ccws.Range(cc_col & ccrow) & vbNewLine 'Program Manager
    End If
    
    ccrow = ccrow + 1 'increment row count in Courtesy Copy for Memos file
    ccstring = ccstring & ccws.Range(cc_col & ccrow) & vbNewLine  'Posting
    
    ccrow = ccrow + 1 'increment row count in Courtesy Copy for Memos file
    ccstring = ccstring & ccws.Range(cc_col & ccrow) & vbNewLine 'File
    
    'insert cc list into mail merge file
    datasourcews.Range("DF2") = ccstring
    
    'Close cc workbook without saving
    ccwb.Close False
    
    'save and close temp mail merge file
    datasourcewb.Close True
            
    'create name for information memo
    smonth = Replace(sample_month, "/", "_")
    'findmemo = "Information Memo for Review Number " & review_number & " for Sample Month " & smonth & ".docm"
    findmemo = "Error Under Threshold Memo for Review Number " & review_number & " for Sample Month " & smonth & ".docx"
    
    'Copy template info memo to active directory
    SrceFile = PathStr & "Finding Memo\Error Under Threshold.docx"
    tempfindmemoname = sPath & "\FM temp.docx"
    DestFile = tempfindmemoname
    FileCopy SrceFile, DestFile
    
    'Check to see if information memo has already been created.
    If Dir(sPath & "\" & findmemo) <> "" Then
        Response = MsgBox("An Information memo has already been created. Do you want to overwrite this memo? No memo will be created if you click NO", Buttons:=vbYesNo)
    ' If statement to check if the NO button was selected.
      If Response = vbNo Then
        End
      End If
    End If

    'set up Word application
    Set appWD = CreateObject("Word.Application")
    appWD.Visible = True

    'open information template
    With appWD
    .Documents.Open fileName:=tempfindmemoname
    
    Application.StatusBar = "Now creating document"

    'merge data from mail merge file with information memo
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
    
    .ActiveDocument.Content.Font.Name = "Arial"
    '.ActiveDocument.Content.Font.Size = 12
    
    'set name for info memo in same folder as schedule
    tempfilename = sPath & "\" & findmemo
    
    
    .ActiveDocument.SaveAs fileName:=tempfilename
                
    .ActiveDocument.Close
    
    'close without saving temp document
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

'tell user that informtaion memo has been created
MsgBox "An Error Under Threshold Memo has been saved as " & sPath & "\" & findmemo & "."

End Sub
Sub MA_SAVE()

'Create Finding Memo
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
SrceFile = PathStr & "Finding Memo/Finding Memo Data Source.xlsx"
DestFile = tempsourcename
FileCopy SrceFile, DestFile

'Open spreadsheet where all the data is stored to populate findings memo
Workbooks.Open fileName:=tempsourcename

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

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'MA

'Client Information
Range("B2") = StrConv(thisws.Range("B2"), vbProperCase) 'Client Name
Range("C2") = Right(thisws.Range("AC10"), 2) & "/" & thisws.Range("H10") 'County and Case Number
Range("D2") = thisws.Range("A10") & thisws.Range("B10") & thisws.Range("C10") & thisws.Range("D10") & thisws.Range("E10") & thisws.Range("F10") 'Review Number
review_number = Range("D2")
Range("E2") = thisws.Range("AM10") 'Review Month
sample_month = Left(thisws.Range("AM10"), 2) & Right(thisws.Range("AM10"), 4)
Range("A7") = Val(thisws.Range("AO3") & thisws.Range("AP3")) 'Reviewer Number
Range("G2") = Range("C7") & "/" & thisws.Range("AG2") 'Supervisor and Reviewer Name
Range("E7") = review_type
Range("F7") = review_type

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

 'area
    areanum = Left(thisws.Range("AC10"), 1)
 'Column C is MA Info column in CC for Memos file
 cc_col = "C"
 
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Going out to website with list of Executive Directors
        FullTextFileName = datasourcews.Range("U6")
            
            'ActiveWorkbook.FollowHyperlink (FullTextFileName)
            Workbooks.Open fileName:=FullTextFileName, ReadOnly:=True
            
            'Allegheny Districts are formatted as dates, change them to 02/XX format
            For irow = 1 To 100
                If IsDate(Range("A" & irow)) Then
                    TempStr = WorksheetFunction.Text(Range("A" & irow), "mm/d/yyyy")
                    Range("A" & irow).NumberFormat = "@"
                    Range("A" & irow) = Left(TempStr, 4)
                End If
            Next irow
            
            For irow = 1 To 1000
                If Range("A" & irow) = countylookup Then
                    datasourcews.Range("Y2") = Range("D" & irow)
                End If
            Next irow
            ActiveWorkbook.Close False
    
'Open CC list file
    Workbooks.Open fileName:=PathStr & "Finding Memo\Courtesy Copy for Memos.xlsx", ReadOnly:=True
    Set ccwb = ActiveWorkbook
    Set ccws = ccwb.Worksheets(areanum)

'Form cc string
    ccstring = vbNewLine & vbNewLine & "cc: "
    ccstring = ccstring & ccws.Range(cc_col & "4") & vbNewLine 'Area Manager
    ccstring = ccstring & ccws.Range(cc_col & "5") & vbNewLine 'Area Staff Asst
    ccrow = 6
    'check area number, if area 1 need to add QA on cc list
    If areanum = "1" Then
        ccstring = ccstring & ccws.Range(cc_col & ccrow) & vbNewLine 'QA for Area 1
        ccrow = ccrow + 1
    End If
    
    'no office managers are listed for info memos
    'ccstring = ccstring & ccws.Range(cc_col & ccrow) & vbNewLine 'Office manager
    ccrow = ccrow + 1 'increment row count in Courtesy Copy for Memos file
    
   'check for commas in Corrective Action names
    If InStr(ccws.Range(cc_col & ccrow), ",") > 0 Then
        'if comma is found, then split string into parts where commas are
        tempArray = Split(ccws.Range(cc_col & ccrow), ",")
        'loop from lower to upper bound of array to get all of the parts
        For isplit = LBound(tempArray) To UBound(tempArray)
            'put each part of string unto a separate line in CC string
           ccstring = ccstring & Trim(tempArray(isplit)) & vbNewLine 'Corrective Action
        Next isplit
    Else 'no commas were found, add whole line into cc string
        ccstring = ccstring & ccws.Range(cc_col & ccrow) & vbNewLine 'Corrective Action
    End If 'check for commas
    ccrow = ccrow + 1 'increment row count in Courtesy Copy for Memos file
    
    'check for dash or blank in Additional Recipients
    If Not (ccws.Range(cc_col & ccrow) = "-" Or ccws.Range(cc_col & ccrow) = "") Then
        'check for commas in Additional Recipient names
        If InStr(ccws.Range(cc_col & ccrow), ",") > 0 Then
            'if comma is found, then split string into parts where commas are
            tempArray = Split(ccws.Range(cc_col & ccrow), ",")
            'loop from lower to upper bound of array to get all of the parts
            For isplit = LBound(tempArray) To UBound(tempArray)
                'put each part of string unto a separate line in CC string
                ccstring = ccstring & Trim(tempArray(isplit)) & vbNewLine 'Additional Recipients
            Next isplit
        Else 'no commas were found, add whole line into cc string
            ccstring = ccstring & ccws.Range(cc_col & ccrow) & vbNewLine 'Additional Recipients
        End If 'check for commas
    End If 'check for dash or blank
    ccrow = ccrow + 1 'increment row count in Courtesy Copy for Memos file
        
    'check for dash or blank in Program Manager, if dash or blank is found, don't add to ccstring
    If Not (ccws.Range(cc_col & ccrow) = "-" Or ccws.Range(cc_col & ccrow) = "") Then
        ccstring = ccstring & ccws.Range(cc_col & ccrow) & vbNewLine 'Program Manager
    End If
    
    ccrow = ccrow + 1 'increment row count in Courtesy Copy for Memos file
    ccstring = ccstring & ccws.Range(cc_col & ccrow) & vbNewLine  'Posting
    
    ccrow = ccrow + 1 'increment row count in Courtesy Copy for Memos file
    ccstring = ccstring & ccws.Range(cc_col & ccrow) & vbNewLine 'File
    
    'insert cc list into mail merge file
    datasourcews.Range("DF2") = ccstring
    
    'Close cc workbook without saving
    ccwb.Close False
    
    'save and close temp mail merge file
    datasourcewb.Close True
            
    'create name for information memo
    smonth = Replace(sample_month, "/", "_")
    'findmemo = "Information Memo for Review Number " & review_number & " for Sample Month " & smonth & ".docm"
    findmemo = "MA SAVE Information Memo for Review Number " & review_number & " for Sample Month " & smonth & ".docx"
    
    'Copy template info memo to active directory
    SrceFile = PathStr & "Finding Memo\MA SAVE Information Memo.docx"
    tempfindmemoname = sPath & "\FM temp.docx"
    DestFile = tempfindmemoname
    FileCopy SrceFile, DestFile
    
    'Check to see if information memo has already been created.
    If Dir(sPath & "\" & findmemo) <> "" Then
        Response = MsgBox("An Information memo has already been created. Do you want to overwrite this memo? No memo will be created if you click NO", Buttons:=vbYesNo)
    ' If statement to check if the NO button was selected.
      If Response = vbNo Then
        End
      End If
    End If

    'set up Word application
    Set appWD = CreateObject("Word.Application")
    appWD.Visible = True

    'open information template
    With appWD
    .Documents.Open fileName:=tempfindmemoname
    
    Application.StatusBar = "Now creating document"

    'merge data from mail merge file with information memo
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
    
    .ActiveDocument.Content.Font.Name = "Arial"
    '.ActiveDocument.Content.Font.Size = 12
    
    'set name for info memo in same folder as schedule
    tempfilename = sPath & "\" & findmemo
    
    
    .ActiveDocument.SaveAs fileName:=tempfilename
                
    .ActiveDocument.Close
    
    'close without saving temp document
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

'tell user that informtaion memo has been created
MsgBox "A MA SAVE Information Memo has been saved as " & sPath & "\" & findmemo & "."

End Sub
Sub Funeral()

'Create Finding Memo
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
SrceFile = PathStr & "Finding Memo/Finding Memo Data Source.xlsx"
DestFile = tempsourcename
FileCopy SrceFile, DestFile

'Open spreadsheet where all the data is stored to populate findings memo
Workbooks.Open fileName:=tempsourcename

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

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'MA

'Client Information
Range("B2") = StrConv(thisws.Range("B2"), vbProperCase) 'Client Name
Range("C2") = Right(thisws.Range("AC10"), 2) & "/" & thisws.Range("H10") 'County and Case Number
Range("D2") = thisws.Range("A10") & thisws.Range("B10") & thisws.Range("C10") & thisws.Range("D10") & thisws.Range("E10") & thisws.Range("F10") 'Review Number
review_number = Range("D2")
Range("E2") = thisws.Range("AM10") 'Review Month
sample_month = Left(thisws.Range("AM10"), 2) & Right(thisws.Range("AM10"), 4)
Range("A7") = Val(thisws.Range("AO3") & thisws.Range("AP3")) 'Reviewer Number
Range("G2") = Range("C7") & "/" & thisws.Range("AG2") 'Supervisor and Reviewer Name
Range("E7") = review_type
Range("F7") = review_type

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

 'area
    areanum = Left(thisws.Range("AC10"), 1)
 'Column C is MA Info column in CC for Memos file
 cc_col = "C"
 
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Going out to website with list of Executive Directors
        FullTextFileName = datasourcews.Range("U6")
            
            'ActiveWorkbook.FollowHyperlink (FullTextFileName)
            Workbooks.Open fileName:=FullTextFileName, ReadOnly:=True
            
            'Allegheny Districts are formatted as dates, change them to 02/XX format
            For irow = 1 To 100
                If IsDate(Range("A" & irow)) Then
                    TempStr = WorksheetFunction.Text(Range("A" & irow), "mm/d/yyyy")
                    Range("A" & irow).NumberFormat = "@"
                    Range("A" & irow) = Left(TempStr, 4)
                End If
            Next irow
            
            For irow = 1 To 1000
                If Range("A" & irow) = countylookup Then
                    datasourcews.Range("Y2") = Range("D" & irow)
                End If
            Next irow
            ActiveWorkbook.Close False
    
'Open CC list file
    Workbooks.Open fileName:=PathStr & "Finding Memo\Courtesy Copy for Memos.xlsx", ReadOnly:=True
    Set ccwb = ActiveWorkbook
    Set ccws = ccwb.Worksheets(areanum)

'Form cc string
    ccstring = vbNewLine & vbNewLine & "cc: "
    ccstring = ccstring & ccws.Range(cc_col & "4") & vbNewLine 'Area Manager
    ccstring = ccstring & ccws.Range(cc_col & "5") & vbNewLine 'Area Staff Asst
    ccrow = 6
    'check area number, if area 1 need to add QA on cc list
    If areanum = "1" Then
        ccstring = ccstring & ccws.Range(cc_col & ccrow) & vbNewLine 'QA for Area 1
        ccrow = ccrow + 1
    End If
    
    'no office managers are listed for info memos
    'ccstring = ccstring & ccws.Range(cc_col & ccrow) & vbNewLine 'Office manager
    ccrow = ccrow + 1 'increment row count in Courtesy Copy for Memos file
    
   'check for commas in Corrective Action names
    If InStr(ccws.Range(cc_col & ccrow), ",") > 0 Then
        'if comma is found, then split string into parts where commas are
        tempArray = Split(ccws.Range(cc_col & ccrow), ",")
        'loop from lower to upper bound of array to get all of the parts
        For isplit = LBound(tempArray) To UBound(tempArray)
            'put each part of string unto a separate line in CC string
           ccstring = ccstring & Trim(tempArray(isplit)) & vbNewLine 'Corrective Action
        Next isplit
    Else 'no commas were found, add whole line into cc string
        ccstring = ccstring & ccws.Range(cc_col & ccrow) & vbNewLine 'Corrective Action
    End If 'check for commas
    ccrow = ccrow + 1 'increment row count in Courtesy Copy for Memos file
    
    'check for dash or blank in Additional Recipients
    If Not (ccws.Range(cc_col & ccrow) = "-" Or ccws.Range(cc_col & ccrow) = "") Then
        'check for commas in Additional Recipient names
        If InStr(ccws.Range(cc_col & ccrow), ",") > 0 Then
            'if comma is found, then split string into parts where commas are
            tempArray = Split(ccws.Range(cc_col & ccrow), ",")
            'loop from lower to upper bound of array to get all of the parts
            For isplit = LBound(tempArray) To UBound(tempArray)
                'put each part of string unto a separate line in CC string
                ccstring = ccstring & Trim(tempArray(isplit)) & vbNewLine 'Additional Recipients
            Next isplit
        Else 'no commas were found, add whole line into cc string
            ccstring = ccstring & ccws.Range(cc_col & ccrow) & vbNewLine 'Additional Recipients
        End If 'check for commas
    End If 'check for dash or blank
    ccrow = ccrow + 1 'increment row count in Courtesy Copy for Memos file
        
    'check for dash or blank in Program Manager, if dash or blank is found, don't add to ccstring
    If Not (ccws.Range(cc_col & ccrow) = "-" Or ccws.Range(cc_col & ccrow) = "") Then
        ccstring = ccstring & ccws.Range(cc_col & ccrow) & vbNewLine 'Program Manager
    End If
    
    ccrow = ccrow + 1 'increment row count in Courtesy Copy for Memos file
    ccstring = ccstring & ccws.Range(cc_col & ccrow) & vbNewLine  'Posting
    
    ccrow = ccrow + 1 'increment row count in Courtesy Copy for Memos file
    ccstring = ccstring & ccws.Range(cc_col & ccrow) & vbNewLine 'File
    
    'insert cc list into mail merge file
    datasourcews.Range("DF2") = ccstring
    
    'Close cc workbook without saving
    ccwb.Close False
    
    'save and close temp mail merge file
    datasourcewb.Close True
            
    'create name for information memo
    smonth = Replace(sample_month, "/", "_")
    'findmemo = "Information Memo for Review Number " & review_number & " for Sample Month " & smonth & ".docm"
    findmemo = "Funeral Home Letter for Review Number " & review_number & " for Sample Month " & smonth & ".docx"
    
    'Copy template info memo to active directory
    SrceFile = PathStr & "Finding Memo\Funeral Home Letter.docx"
    tempfindmemoname = sPath & "\FM temp.docx"
    DestFile = tempfindmemoname
    FileCopy SrceFile, DestFile
    
    'Check to see if information memo has already been created.
    If Dir(sPath & "\" & findmemo) <> "" Then
        Response = MsgBox("A Funeral Home Letter has already been created. Do you want to overwrite this memo? No memo will be created if you click NO", Buttons:=vbYesNo)
    ' If statement to check if the NO button was selected.
      If Response = vbNo Then
        End
      End If
    End If

    'set up Word application
    Set appWD = CreateObject("Word.Application")
    appWD.Visible = True

    'open information template
    With appWD
    .Documents.Open fileName:=tempfindmemoname
    
    Application.StatusBar = "Now creating document"

    'merge data from mail merge file with information memo
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
    
    .ActiveDocument.Content.Font.Name = "Arial"
    '.ActiveDocument.Content.Font.Size = 12
    
    'set name for info memo in same folder as schedule
    tempfilename = sPath & "\" & findmemo
    
    
    .ActiveDocument.SaveAs fileName:=tempfilename
                
    .ActiveDocument.Close
    
    'close without saving temp document
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

'tell user that informtaion memo has been created
MsgBox "A Funeral Home Letter has been saved as " & sPath & "\" & findmemo & "."

End Sub
Sub Adoption()

'Create Finding Memo
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
SrceFile = PathStr & "Finding Memo/Finding Memo Data Source.xlsx"
DestFile = tempsourcename
FileCopy SrceFile, DestFile

'Open spreadsheet where all the data is stored to populate findings memo
Workbooks.Open fileName:=tempsourcename

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

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'MA

'Client Information
Range("B2") = StrConv(thisws.Range("B2"), vbProperCase) 'Client Name
Range("C2") = Right(thisws.Range("AC10"), 2) & "/" & thisws.Range("H10") 'County and Case Number
Range("D2") = thisws.Range("A10") & thisws.Range("B10") & thisws.Range("C10") & thisws.Range("D10") & thisws.Range("E10") & thisws.Range("F10") 'Review Number
review_number = Range("D2")
Range("E2") = thisws.Range("AM10") 'Review Month
sample_month = Left(thisws.Range("AM10"), 2) & Right(thisws.Range("AM10"), 4)
Range("A7") = Val(thisws.Range("AO3") & thisws.Range("AP3")) 'Reviewer Number
Range("G2") = Range("C7") & "/" & thisws.Range("AG2") 'Supervisor and Reviewer Name
Range("E7") = review_type
Range("F7") = review_type

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

 'area
    areanum = Left(thisws.Range("AC10"), 1)
 'Column C is MA Info column in CC for Memos file
 cc_col = "C"
 
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Going out to website with list of Executive Directors
        FullTextFileName = datasourcews.Range("U6")
            
            'ActiveWorkbook.FollowHyperlink (FullTextFileName)
            Workbooks.Open fileName:=FullTextFileName, ReadOnly:=True
            
            'Allegheny Districts are formatted as dates, change them to 02/XX format
            For irow = 1 To 100
                If IsDate(Range("A" & irow)) Then
                    TempStr = WorksheetFunction.Text(Range("A" & irow), "mm/d/yyyy")
                    Range("A" & irow).NumberFormat = "@"
                    Range("A" & irow) = Left(TempStr, 4)
                End If
            Next irow
            
            For irow = 1 To 1000
                If Range("A" & irow) = countylookup Then
                    datasourcews.Range("Y2") = Range("D" & irow)
                End If
            Next irow
            ActiveWorkbook.Close False
    
'Open CC list file
    Workbooks.Open fileName:=PathStr & "Finding Memo\Courtesy Copy for Memos.xlsx", ReadOnly:=True
    Set ccwb = ActiveWorkbook
    Set ccws = ccwb.Worksheets(areanum)

'Form cc string
    ccstring = vbNewLine & vbNewLine & "cc: "
    ccstring = ccstring & ccws.Range(cc_col & "4") & vbNewLine 'Area Manager
    ccstring = ccstring & ccws.Range(cc_col & "5") & vbNewLine 'Area Staff Asst
    ccrow = 6
    'check area number, if area 1 need to add QA on cc list
    If areanum = "1" Then
        ccstring = ccstring & ccws.Range(cc_col & ccrow) & vbNewLine 'QA for Area 1
        ccrow = ccrow + 1
    End If
    
    'no office managers are listed for info memos
    'ccstring = ccstring & ccws.Range(cc_col & ccrow) & vbNewLine 'Office manager
    ccrow = ccrow + 1 'increment row count in Courtesy Copy for Memos file
    
   'check for commas in Corrective Action names
    If InStr(ccws.Range(cc_col & ccrow), ",") > 0 Then
        'if comma is found, then split string into parts where commas are
        tempArray = Split(ccws.Range(cc_col & ccrow), ",")
        'loop from lower to upper bound of array to get all of the parts
        For isplit = LBound(tempArray) To UBound(tempArray)
            'put each part of string unto a separate line in CC string
           ccstring = ccstring & Trim(tempArray(isplit)) & vbNewLine 'Corrective Action
        Next isplit
    Else 'no commas were found, add whole line into cc string
        ccstring = ccstring & ccws.Range(cc_col & ccrow) & vbNewLine 'Corrective Action
    End If 'check for commas
    ccrow = ccrow + 1 'increment row count in Courtesy Copy for Memos file
    
    'check for dash or blank in Additional Recipients
    If Not (ccws.Range(cc_col & ccrow) = "-" Or ccws.Range(cc_col & ccrow) = "") Then
        'check for commas in Additional Recipient names
        If InStr(ccws.Range(cc_col & ccrow), ",") > 0 Then
            'if comma is found, then split string into parts where commas are
            tempArray = Split(ccws.Range(cc_col & ccrow), ",")
            'loop from lower to upper bound of array to get all of the parts
            For isplit = LBound(tempArray) To UBound(tempArray)
                'put each part of string unto a separate line in CC string
                ccstring = ccstring & Trim(tempArray(isplit)) & vbNewLine 'Additional Recipients
            Next isplit
        Else 'no commas were found, add whole line into cc string
            ccstring = ccstring & ccws.Range(cc_col & ccrow) & vbNewLine 'Additional Recipients
        End If 'check for commas
    End If 'check for dash or blank
    ccrow = ccrow + 1 'increment row count in Courtesy Copy for Memos file
        
    'check for dash or blank in Program Manager, if dash or blank is found, don't add to ccstring
    If Not (ccws.Range(cc_col & ccrow) = "-" Or ccws.Range(cc_col & ccrow) = "") Then
        ccstring = ccstring & ccws.Range(cc_col & ccrow) & vbNewLine 'Program Manager
    End If
    
    ccrow = ccrow + 1 'increment row count in Courtesy Copy for Memos file
    ccstring = ccstring & ccws.Range(cc_col & ccrow) & vbNewLine  'Posting
    
    ccrow = ccrow + 1 'increment row count in Courtesy Copy for Memos file
    ccstring = ccstring & ccws.Range(cc_col & ccrow) & vbNewLine 'File
    
    'insert cc list into mail merge file
    datasourcews.Range("DF2") = ccstring
    
    'Close cc workbook without saving
    ccwb.Close False
    
    'save and close temp mail merge file
    datasourcewb.Close True
            
    'create name for information memo
    smonth = Replace(sample_month, "/", "_")
    'findmemo = "Information Memo for Review Number " & review_number & " for Sample Month " & smonth & ".docm"
    findmemo = "MA Adoption Memo for Review Number " & review_number & " for Sample Month " & smonth & ".docx"
    
    'Copy template info memo to active directory
    SrceFile = PathStr & "Finding Memo\MA Adoption Foster Care.docx"
    tempfindmemoname = sPath & "\FM temp.docx"
    DestFile = tempfindmemoname
    FileCopy SrceFile, DestFile
    
    'Check to see if information memo has already been created.
    If Dir(sPath & "\" & findmemo) <> "" Then
        Response = MsgBox("An Information memo has already been created. Do you want to overwrite this memo? No memo will be created if you click NO", Buttons:=vbYesNo)
    ' If statement to check if the NO button was selected.
      If Response = vbNo Then
        End
      End If
    End If

    'set up Word application
    Set appWD = CreateObject("Word.Application")
    appWD.Visible = True

    'open information template
    With appWD
    .Documents.Open fileName:=tempfindmemoname
    
    Application.StatusBar = "Now creating document"

    'merge data from mail merge file with information memo
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
    
    .ActiveDocument.Content.Font.Name = "Arial"
    '.ActiveDocument.Content.Font.Size = 12
    
    'set name for info memo in same folder as schedule
    tempfilename = sPath & "\" & findmemo
    
    
    .ActiveDocument.SaveAs fileName:=tempfilename
                
    .ActiveDocument.Close
    
    'close without saving temp document
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

'tell user that informtaion memo has been created
MsgBox "A MA Adoption Memo has been saved as " & sPath & "\" & findmemo & "."

End Sub

Sub TANF_SAVE()

'Create Finding Memo
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
SrceFile = PathStr & "Finding Memo/Finding Memo Data Source.xlsx"
DestFile = tempsourcename
FileCopy SrceFile, DestFile

'Open spreadsheet where all the data is stored to populate findings memo
Workbooks.Open fileName:=tempsourcename

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

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'TANF

' Get county number and name in correct format
countynum = Right(thisws.Range("U10"), 2)
districtnum = thisws.Range("Y10")
countylookup = countyformat(countynum, thisws.Range("O4"), districtnum)
Range("AK2") = thisws.Range("O4") & " Assistance Office"
'Look for counties with districts
'If countynum = "02" Or countynum = "23" Or countynum = "40" Or _
'        countynum = "51" Or countynum = "63" Or countynum = "65" Then
'    Range("AK2") = thisws.Range("O4") & " Office"
'Else
'    Range("AK2") = thisws.Range("O4") & " CAO"
'End If

'Client Information
Range("B2") = StrConv(thisws.Range("B2"), vbProperCase) 'Client Name
Range("C2") = Right(thisws.Range("U10"), 2) & "/" & thisws.Range("I10") 'County and Case Number
Range("D2") = thisws.Range("A10") 'Review Number
review_number = Range("D2")
Range("E2") = Left(thisws.Range("AB10"), 2) & "/" & Right(thisws.Range("AB10"), 4) 'Review Month
sample_month = thisws.Range("AB10")
Range("A7") = Val(thisws.Range("AO3") & thisws.Range("AP3")) 'Reviewer Number
Range("G2") = Range("C7") & "/" & thisws.Range("AG2") 'Supervisor and Reviewer Name
Range("E7") = review_type
Range("F7") = review_type
Range("R2") = thisws.Range("J61") & " - " & thisws.Range("O61") & " - " & thisws.Range("T61") 'Element, Nature, & Cause Code
Range("E17") = Val(thisws.Range("AL10")) 'Review Findings
If review_type = 1 Then
Range("O2") = thiswb.Sheets("TANF Computation").Range("B72") 'Benefit Amt
Range("P2") = thiswb.Sheets("TANF Computation").Range("C69") 'Benefit Determination
Else
Range("O2") = thiswb.Sheets("GA Computation").Range("B72") 'Benefit Amt
Range("P2") = thiswb.Sheets("GA Computation").Range("C69") 'Benefit Determination
End If
Range("Q2") = thisws.Range("AO10") 'Amt of Error

'reformat street address
    Range("Z2") = StrConv(thisws.Range("B3"), vbProperCase)
'Reformat city and state
If review_type = 1 Then
    temparry = Split(thisws.Range("B4"), ",")
    TempStr = StrConv(Trim(temparry(LBound(temparry))), vbProperCase)
    Range("AA2") = TempStr & ", " & Trim(temparry(UBound(temparry)))
Else
     temparry = Split(thisws.Range("B5"), ",")
    TempStr = StrConv(Trim(temparry(LBound(temparry))), vbProperCase)
    Range("AA2") = TempStr & ", " & Trim(temparry(UBound(temparry)))
End If
'category and grant group
Range("CE2") = thisws.Range("Q10")
Range("CG2") = thisws.Range("S10")
'area
areanum = Left(thisws.Range("U10"), 1)
'Column G is TANF Info column in Courtesy Copy for Memos file
cc_col = "G"

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Going out to website with list of Executive Directors
        FullTextFileName = datasourcews.Range("U6")
            
            'ActiveWorkbook.FollowHyperlink (FullTextFileName)
            Workbooks.Open fileName:=FullTextFileName, ReadOnly:=True
            
            'Allegheny Districts are formatted as dates, change them to 02/XX format
            For irow = 1 To 100
                If IsDate(Range("A" & irow)) Then
                    TempStr = WorksheetFunction.Text(Range("A" & irow), "mm/d/yyyy")
                    Range("A" & irow).NumberFormat = "@"
                    Range("A" & irow) = Left(TempStr, 4)
                End If
            Next irow
            
            For irow = 1 To 1000
                If Range("A" & irow) = countylookup Then
                    datasourcews.Range("Y2") = Range("D" & irow)
                End If
            Next irow
            ActiveWorkbook.Close False
    
'Open CC list file
    Workbooks.Open fileName:=PathStr & "Finding Memo\Courtesy Copy for Memos.xlsx", ReadOnly:=True
    Set ccwb = ActiveWorkbook
    Set ccws = ccwb.Worksheets(areanum)

'Form cc string
    ccstring = vbNewLine & vbNewLine & "cc: "
    ccstring = ccstring & ccws.Range(cc_col & "4") & vbNewLine 'Area Manager
    ccstring = ccstring & ccws.Range(cc_col & "5") & vbNewLine 'Area Staff Asst
    ccrow = 6
    'check area number, if area 1 need to add QA on cc list
    If areanum = "1" Then
        ccstring = ccstring & ccws.Range(cc_col & ccrow) & vbNewLine 'QA for Area 1
        ccrow = ccrow + 1
    End If
    
    'no office managers are listed for info memos
    'ccstring = ccstring & ccws.Range(cc_col & ccrow) & vbNewLine 'Office manager
    ccrow = ccrow + 1 'increment row count in Courtesy Copy for Memos file
    
   'check for commas in Corrective Action names
    If InStr(ccws.Range(cc_col & ccrow), ",") > 0 Then
        'if comma is found, then split string into parts where commas are
        tempArray = Split(ccws.Range(cc_col & ccrow), ",")
        'loop from lower to upper bound of array to get all of the parts
        For isplit = LBound(tempArray) To UBound(tempArray)
            'put each part of string unto a separate line in CC string
           ccstring = ccstring & Trim(tempArray(isplit)) & vbNewLine 'Corrective Action
        Next isplit
    Else 'no commas were found, add whole line into cc string
        ccstring = ccstring & ccws.Range(cc_col & ccrow) & vbNewLine 'Corrective Action
    End If 'check for commas
    ccrow = ccrow + 1 'increment row count in Courtesy Copy for Memos file
    
    'check for dash or blank in Additional Recipients
    If Not (ccws.Range(cc_col & ccrow) = "-" Or ccws.Range(cc_col & ccrow) = "") Then
        'check for commas in Additional Recipient names
        If InStr(ccws.Range(cc_col & ccrow), ",") > 0 Then
            'if comma is found, then split string into parts where commas are
            tempArray = Split(ccws.Range(cc_col & ccrow), ",")
            'loop from lower to upper bound of array to get all of the parts
            For isplit = LBound(tempArray) To UBound(tempArray)
                'put each part of string unto a separate line in CC string
                ccstring = ccstring & Trim(tempArray(isplit)) & vbNewLine 'Additional Recipients
            Next isplit
        Else 'no commas were found, add whole line into cc string
            ccstring = ccstring & ccws.Range(cc_col & ccrow) & vbNewLine 'Additional Recipients
        End If 'check for commas
    End If 'check for dash or blank
    ccrow = ccrow + 1 'increment row count in Courtesy Copy for Memos file
        
    'check for dash or blank in Program Manager, if dash or blank is found, don't add to ccstring
    If Not (ccws.Range(cc_col & ccrow) = "-" Or ccws.Range(cc_col & ccrow) = "") Then
        ccstring = ccstring & ccws.Range(cc_col & ccrow) & vbNewLine 'Program Manager
    End If
    
    ccrow = ccrow + 1 'increment row count in Courtesy Copy for Memos file
    ccstring = ccstring & ccws.Range(cc_col & ccrow) & vbNewLine  'Posting
    
    ccrow = ccrow + 1 'increment row count in Courtesy Copy for Memos file
    ccstring = ccstring & ccws.Range(cc_col & ccrow) & vbNewLine 'File
    
    'insert cc list into mail merge file
    datasourcews.Range("DF2") = ccstring
    
    'Close cc workbook without saving
    ccwb.Close False
    
    'save and close temp mail merge file
    datasourcewb.Close True
            
    'create name for information memo
    smonth = Replace(sample_month, "/", "_")
    'findmemo = "Information Memo for Review Number " & review_number & " for Sample Month " & smonth & ".docm"
    If review_type = 1 Then
    findmemo = "TANF SAVE Information Memo for Review Number " & review_number & " for Sample Month " & smonth & ".docx"
    Else
    findmemo = "GA SAVE Information Memo for Review Number " & review_number & " for Sample Month " & smonth & ".docx"
    End If
    
    'Copy template info memo to active directory
    SrceFile = PathStr & "Finding Memo\TANF SAVE Information Memo.docx"
    tempfindmemoname = sPath & "\FM temp.docx"
    DestFile = tempfindmemoname
    FileCopy SrceFile, DestFile
    
    'Check to see if information memo has already been created.
    If Dir(sPath & "\" & findmemo) <> "" Then
        Response = MsgBox("An Information memo has already been created. Do you want to overwrite this memo? No memo will be created if you click NO", Buttons:=vbYesNo)
    ' If statement to check if the NO button was selected.
      If Response = vbNo Then
        End
      End If
    End If

    'set up Word application
    Set appWD = CreateObject("Word.Application")
    appWD.Visible = True

    'open information template
    With appWD
    .Documents.Open fileName:=tempfindmemoname
    
    Application.StatusBar = "Now creating document"

    'merge data from mail merge file with information memo
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
    
    .ActiveDocument.Content.Font.Name = "Arial"
    '.ActiveDocument.Content.Font.Size = 12
    
    'set name for info memo in same folder as schedule
    tempfilename = sPath & "\" & findmemo
    
    
    .ActiveDocument.SaveAs fileName:=tempfilename
                
    .ActiveDocument.Close
    
    'close without saving temp document
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

'tell user that informtaion memo has been created
If review_type = 1 Then
MsgBox "A TANF SAVE Information Memo has been saved as " & sPath & "\" & findmemo & "."
Else
MsgBox "A GA SAVE Information Memo has been saved as " & sPath & "\" & findmemo & "."
End If
End Sub

Sub LEP()

'Create Finding Memo
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
SrceFile = PathStr & "Finding Memo/Finding Memo Data Source.xlsx"
DestFile = tempsourcename
FileCopy SrceFile, DestFile

'Open spreadsheet where all the data is stored to populate findings memo
Workbooks.Open fileName:=tempsourcename

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

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'TANF
If review_type = 1 Then

' Get county number and name in correct format
countynum = Right(thisws.Range("U10"), 2)
districtnum = thisws.Range("Y10")
countylookup = countyformat(countynum, thisws.Range("O4"), districtnum)
Range("AK2") = thisws.Range("O4") & " Assistance Office"


'Client Information
Range("B2") = StrConv(thisws.Range("B2"), vbProperCase) 'Client Name
Range("C2") = Right(thisws.Range("U10"), 2) & "/" & thisws.Range("I10") 'County and Case Number
Range("D2") = thisws.Range("A10") 'Review Number
review_number = Range("D2")
Range("E2") = Left(thisws.Range("AB10"), 2) & "/" & Right(thisws.Range("AB10"), 4) 'Review Month
sample_month = thisws.Range("AB10")
Range("A7") = Val(thisws.Range("AO3") & thisws.Range("AP3")) 'Reviewer Number
Range("G2") = Range("C7") & "/" & thisws.Range("AG2") 'Supervisor and Reviewer Name
Range("E7") = review_type
Range("F7") = review_type
Range("R2") = thisws.Range("J61") & " - " & thisws.Range("O61") & " - " & thisws.Range("T61") 'Element, Nature, & Cause Code
Range("E17") = Val(thisws.Range("AL10")) 'Review Findings
Range("O2") = thiswb.Sheets("TANF Computation").Range("B72") 'Benefit Amt
Range("P2") = thiswb.Sheets("TANF Computation").Range("C69") 'Benefit Determination
Range("Q2") = thisws.Range("AO10") 'Amt of Error

'reformat street address
    Range("Z2") = StrConv(thisws.Range("B3"), vbProperCase)
'Reformat city and state
    temparry = Split(thisws.Range("B4"), ",")
    TempStr = StrConv(Trim(temparry(LBound(temparry))), vbProperCase)
    Range("AA2") = TempStr & ", " & Trim(temparry(UBound(temparry)))
    
'category and grant group
Range("CE2") = thisws.Range("Q10")
Range("CG2") = thisws.Range("S10")
'area
areanum = Left(thisws.Range("U10"), 1)
'Column G is TANF Info column in Courtesy Copy for Memos file
cc_col = "G"

'/////////////////////////////////////////////////////////////////////////////////////////////////////////////

'SNAP Positive
ElseIf review_type = 5 Then

'Client Information
Range("B2") = StrConv(thisws.Range("B4"), vbProperCase)
Range("C2") = Right(thisws.Range("X18"), 2) & "/" & Left(thisws.Range("I18"), 9)

' Get county number and name in correct format
countynum = Right(thisws.Range("X18"), 2)
districtnum = thisws.Range("B153") & thisws.Range("C153")
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
'category
Range("CE2") = "FS"
'area
areanum = Left(thisws.Range("X18"), 1)
'Column E is SNAP Info column in Courtesy Copy for Memos file
cc_col = "E"

'///////////////////////////////////////////////////////////////////////////////////////////////////
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

areanum = Left(thisws.Range("AA20"), 1)
'Column E is SNAP Info column in Courtesy Copy for Memos file
cc_col = "E"


'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

'MA Positive
ElseIf review_type = 2 Then

' Get county number and name in correct format
countynum = Right(thisws.Range("AC10"), 2)
districtnum = thisws.Range("AH10")
countylookup = countyformat(countynum, thisws.Range("O4"), districtnum)
Range("AK2") = thisws.Range("O4") & " Assistance Office"


'Client Information
Range("B2") = StrConv(thisws.Range("B2"), vbProperCase) 'Client Name
Range("C2") = Right(thisws.Range("AC10"), 2) & "/" & thisws.Range("H10") 'County and Case Number
Range("D2") = thisws.Range("A10") & thisws.Range("B10") & thisws.Range("C10") & thisws.Range("D10") & thisws.Range("E10") & thisws.Range("F10") 'Review Number
review_number = Range("D2")
Range("E2") = thisws.Range("AM10") 'Review Month
sample_month = thisws.Range("AM10")
Range("A7") = Val(thisws.Range("AO3") & thisws.Range("AP3")) 'Reviewer Number
Range("G2") = Range("C7") & "/" & thisws.Range("AG2") 'Supervisor and Reviewer Name
Range("E7") = review_type
Range("F7") = review_type

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
'category and grant group
Range("CE2") = thisws.Range("Q10")
Range("CG2") = thisws.Range("Z10")
'area
areanum = Left(thisws.Range("AC10"), 1)
'Column C is MA Info column in Courtesy Copy for Memos file
cc_col = "C"

'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'MA Negative
ElseIf review_type = 8 Then

' Get county number and name in correct format
countynum = Right(thisws.Range("G15"), 2)
districtnum = thisws.Range("K15")
countylookup = countyformat(countynum, thisws.Range("F2"), districtnum)
Range("AK2") = thisws.Range("F2") & " Assistance Office"

'Client Information
Range("B2") = StrConv(thisws.Range("B7"), vbProperCase)
'reformat street address
If thisws.Range("L7") <> "" Then
    Range("Z2") = StrConv(thisws.Range("L6"), vbProperCase) & ", " & StrConv(thisws.Range("L7"), vbProperCase)
Else
    Range("Z2") = StrConv(thisws.Range("L6"), vbProperCase)
End If

'Reformat city and state
    temparry = Split(thisws.Range("L8"), ",")
    TempStr = StrConv(Trim(temparry(LBound(temparry))), vbProperCase)

Range("AA2") = TempStr & ", " & Trim(temparry(UBound(temparry)))

'Case Number
Range("C2") = Right(thisws.Range("G15"), 2) & "/" & thisws.Range("S15")

'review number
Range("D2") = thisws.Range("L15")
review_number = Range("D2")
'review month
Range("E2") = thisws.Range("T11")
sample_month = thisws.Range("T11")
'reviewer number
Range("A7") = Val(thisws.Range("AB11") & thisws.Range("AC11"))
'supervisor/examiner
Range("G2") = Range("C7") & "/" & Range("BA2")
'Range("R2") = thisws.Range("J61") & " - " & thisws.Range("O61")
'finding
Range("E17") = Val(thisws.Range("AL10"))
Range("F7") = review_type
Range("E7") = review_type
'category and grant group
Range("CE2") = thisws.Range("AB15")
Range("CG2") = thisws.Range("F19")
'area
areanum = Left(thisws.Range("G15"), 1)
'Column C is MA Info column in Courtesy Copy for Memos file
cc_col = "C"

End If
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Going out to website with list of Executive Directors
        FullTextFileName = datasourcews.Range("U6")
            
            'ActiveWorkbook.FollowHyperlink (FullTextFileName)
            Workbooks.Open fileName:=FullTextFileName, ReadOnly:=True
            
            'Allegheny Districts are formatted as dates, change them to 02/XX format
            For irow = 1 To 100
                If IsDate(Range("A" & irow)) Then
                    TempStr = WorksheetFunction.Text(Range("A" & irow), "mm/d/yyyy")
                    Range("A" & irow).NumberFormat = "@"
                    Range("A" & irow) = Left(TempStr, 4)
                End If
            Next irow
            
            For irow = 1 To 1000
                If Range("A" & irow) = countylookup Then
                    datasourcews.Range("Y2") = Range("D" & irow)
                End If
            Next irow
            ActiveWorkbook.Close False
            
    
'Open CC list file
    Workbooks.Open fileName:=PathStr & "Finding Memo\Courtesy Copy for Memos.xlsx", ReadOnly:=True
    Set ccwb = ActiveWorkbook
    Set ccws = ccwb.Worksheets(areanum)

'Form cc string to add to memo at bottom
    ccstring = vbNewLine & vbNewLine & "cc: " 'first line has cc: then add new line to string
    ccstring = ccstring & ccws.Range(cc_col & "4") & vbNewLine 'Area Manager
    ccstring = ccstring & ccws.Range(cc_col & "5") & vbNewLine 'Area Staff Asst
    'create variable to keep track of which row we are accessing in the Courtesy Copy for Memos file
    'need this variable because Area 1 has an extra line for their QA person
    ccrow = 6
    'check area number, if area 1 need to add QA on cc list
    If areanum = "1" Then
        ccstring = ccstring & ccws.Range(cc_col & ccrow) & vbNewLine 'QA for Area 1
        ccrow = ccrow + 1 'increment row count in Courtesy Copy for Memos file
    End If
    
    'no office managers are listed for info memos
    'ccstring = ccstring & ccws.Range(cc_col & ccrow) & vbNewLine 'Office manager
    ccrow = ccrow + 1 'skip over office manager but need to increment row count
    
   'check for commas in Corrective Action names
    If InStr(ccws.Range(cc_col & ccrow), ",") > 0 Then
        'if comma is found, then split string into parts where commas are
        tempArray = Split(ccws.Range(cc_col & ccrow), ",")
        'loop from lower to upper bound of array to get all of the parts
        For isplit = LBound(tempArray) To UBound(tempArray)
            'put each part of string unto a separate line in CC string
           ccstring = ccstring & Trim(tempArray(isplit)) & vbNewLine 'Corrective Action
        Next isplit
    Else 'no commas were found, add whole line into cc string
        ccstring = ccstring & ccws.Range(cc_col & ccrow) & vbNewLine 'Corrective Action
    End If 'check for commas
    ccrow = ccrow + 1 'increment row count in Courtesy Copy for Memos file
    
    'check for dash or blank in Additional Recipients
    If Not (ccws.Range(cc_col & ccrow) = "-" Or ccws.Range(cc_col & ccrow) = "") Then
        'check for commas in Additional Recipient names
        If InStr(ccws.Range(cc_col & ccrow), ",") > 0 Then
            'if comma is found, then split string into parts where commas are
            tempArray = Split(ccws.Range(cc_col & ccrow), ",")
            'loop from lower to upper bound of array to get all of the parts
            For isplit = LBound(tempArray) To UBound(tempArray)
                'put each part of string unto a separate line in CC string
                ccstring = ccstring & Trim(tempArray(isplit)) & vbNewLine 'Additional Recipients
            Next isplit
        Else 'no commas were found, add whole line into cc string
            ccstring = ccstring & ccws.Range(cc_col & ccrow) & vbNewLine 'Additional Recipients
        End If 'check for commas
    End If 'check for dash or blank
    ccrow = ccrow + 1 'increment row count in Courtesy Copy for Memos file
        
    'check for dash or blank in Program Manager, if dash or blank is found, don't add to ccstring
    If Not (ccws.Range(cc_col & ccrow) = "-" Or ccws.Range(cc_col & ccrow) = "") Then
        ccstring = ccstring & ccws.Range(cc_col & ccrow) & vbNewLine 'Program Manager
    End If
    
    ccrow = ccrow + 1 'increment row count in Courtesy Copy for Memos file
    ccstring = ccstring & ccws.Range(cc_col & ccrow) & vbNewLine  'Posting
    
    ccrow = ccrow + 1 'increment row count in Courtesy Copy for Memos file
    ccstring = ccstring & ccws.Range(cc_col & ccrow) & vbNewLine 'File
    
    'insert cc list into mail merge file
    datasourcews.Range("DF2") = ccstring
    
    'Close cc workbook without saving
    ccwb.Close False
    
    'save and close temp mail merge file
    datasourcewb.Close True
            
    'create name for information memo
    smonth = Replace(sample_month, "/", "_")
    'findmemo = "Information Memo for Review Number " & review_number & " for Sample Month " & smonth & ".docm"
    findmemo = "LEP Information Memo for Review Number " & review_number & " for Sample Month " & smonth & ".docx"
    
    'Copy template info memo to active directory
    SrceFile = PathStr & "Finding Memo\LEP Memo.docx"
    tempfindmemoname = sPath & "\FM temp.docx"
    DestFile = tempfindmemoname
    FileCopy SrceFile, DestFile
    
    'Check to see if information memo has already been created.
    If Dir(sPath & "\" & findmemo) <> "" Then
        Response = MsgBox("A LEP Info. memo has already been created. Do you want to overwrite this memo? No memo will be created if you click NO", Buttons:=vbYesNo)
    ' If statement to check if the NO button was selected.
      If Response = vbNo Then
        End
      End If
    End If

    'set up Word application
    Set appWD = CreateObject("Word.Application")
    appWD.Visible = True

    'open information template
    With appWD
    .Documents.Open fileName:=tempfindmemoname
    
    Application.StatusBar = "Now creating document"

    'merge data from mail merge file with information memo
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
    
    .ActiveDocument.Content.Font.Name = "Arial"
    '.ActiveDocument.Content.Font.Size = 12
    
    'set name for info memo in same folder as schedule
    tempfilename = sPath & "\" & findmemo
    
    '******************************************************************************************
    '******************************************************************************************
    'This section was to add a button to information memo to copy memo to paperless folder
    'but doesn't seem to work consistently
    '******************************************************************************************
    
    'add a button to the end of the info memo to copy info memo to paperless folder
    'Dim oRng As Word.Range
    '.ActiveDocument.Range.InsertAfter vbCr
    'Set oRng = .ActiveDocument.Range
    'Dim shp As Word.InlineShape
    'oRng.Collapse wdCollapseEnd
    'put button at end
    'Set shp = oRng.InlineShapes.AddOLEControl(ClassType:="Forms.CommandButton.1")
    'make button larger with word wrap
    'shp.OLEFormat.Object.WordWrap = True
    'shp.OLEFormat.Object.FontSize = 10
    'shp.OLEFormat.Object.Height = 56
    'shp.OLEFormat.Object.Width = 56
    'shp.OLEFormat.Object.Caption = "Copy to Paperless"
    
    'Add the procedure to saveas to the paperless folder on the click event of the button
    'change the pathpaperless variable below to save to a different folder
    'Dim sCode As String
    'sCode = "Private Sub " & shp.OLEFormat.Object.Name & "_Click()" & vbCrLf & _
    '        "'Save current document" & vbCrLf & _
    '        "   ActiveDocument.Save" & vbCrLf & _
    '        "'Gets name of document" & vbCrLf & _
    '        "   docname=ActiveDocument.Name" & vbCrLf & _
    '        "'Changes file ending from macro to non-macro document" & vbCrLf & _
    '        "   docname=Replace(docname, "".docm"", "".docx"")" & vbCrLf & _
    '        "'Set path to folder" & vbCrLf & _
    '        "   pathpaperless = ""\\dhs\share\oim\pwimdaubts04\data\stat\dqc\Paperless & Speechless\""" & vbCrLf & _
    '        "   If Dir(pathpaperless, vbDirectory) = """" Then" & vbCrLf & _
    '        "      MsgBox ""The folder Paperless is not on your computer. Stopping program.""" & vbCrLf & _
    '        "      End" & vbCrLf & _
    '        "   End If" & vbCrLf & _
    '        "'Creates new document" & vbCrLf & _
    '        "   Application.Documents.Add ActiveDocument.FullName" & vbCrLf & _
    '        "'Deletes the button on document" & vbCrLf & _
    '        "   For Each o In ActiveDocument.InlineShapes" & vbCrLf & _
    '        "      If o.OLEFormat.Object.Caption = ""Copy to Paperless"" Then" & vbCrLf & _
    '        "          o.Delete" & vbCrLf & _
    '        "      End If" & vbCrLf & _
    '        "   Next" & vbCrLf & _
    '        "'Saves document to desired folder" & vbCrLf & _
    '        "   ActiveDocument.SaveAs2 FileName:=pathpaperless & docname, FileFormat:= wdFormatXMLDocument" & vbCrLf & _
    '        "   ActiveDocument.Close" & vbCrLf & _
    '        "End Sub"
    '.ActiveDocument.VBProject.VBComponents("ThisDocument").CodeModule.AddFromString sCode

    'save information memo with name in macro format
    '.ActiveDocument.SaveAs2 FileName:=tempfilename, _
    '    FileFormat:=wdFormatXMLDocumentMacroEnabled, _
    '    LockComments:=False, _
    '    Password:="", AddToRecentFiles:=True, WritePassword:="", _
    '    ReadOnlyRecommended:=False, EmbedTrueTypeFonts:=False, _
    '    SaveNativePictureFormat:=False, SaveFormsData:=False, SaveAsAOCELetter:= _
    '    False, CompatibilityMode:=14
        
    '    tempmodname = "Module1"
    
    '******************************************************************************************
    '******************************************************************************************
    
    .ActiveDocument.SaveAs fileName:=tempfilename
                
    .ActiveDocument.Close
    
    'close without saving temp document
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

'tell user that informtaion memo has been created
MsgBox "An Information Memo has been saved as " & sPath & "\" & findmemo & "."

End Sub

Sub TANFPend()

'Create Finding Memo
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

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'TANF
If review_type = 1 Then

' Get county number and name in correct format
countynum = Right(thisws.Range("U10"), 2)
districtnum = thisws.Range("Y10")
countylookup = countyformat(countynum, thisws.Range("O4"), districtnum)
Range("AK2") = thisws.Range("O4") & " Assistance Office"
'Look for counties with districts
'If countynum = "02" Or countynum = "23" Or countynum = "40" Or _
'        countynum = "51" Or countynum = "63" Or countynum = "65" Then
'    Range("AK2") = thisws.Range("O4") & " Office"
'Else
'    Range("AK2") = thisws.Range("O4") & " CAO"
'End If

'Client Information
Range("B2") = StrConv(thisws.Range("B2"), vbProperCase) 'Client Name
Range("C2") = Right(thisws.Range("U10"), 2) & "/" & thisws.Range("I10") 'County and Case Number
Range("D2") = thisws.Range("A10") 'Review Number
review_number = Range("D2")
Range("E2") = Left(thisws.Range("AB10"), 2) & "/" & Right(thisws.Range("AB10"), 4) 'Review Month
sample_month = thisws.Range("AB10")
Range("A7") = Val(thisws.Range("AO3") & thisws.Range("AP3")) 'Reviewer Number
Range("G2") = Range("C7") & "/" & thisws.Range("AG2") 'Supervisor and Reviewer Name
Range("E7") = review_type
Range("F7") = review_type
Range("R2") = thisws.Range("J61") & " - " & thisws.Range("O61") & " - " & thisws.Range("T61") 'Element, Nature, & Cause Code
Range("E17") = Val(thisws.Range("AL10")) 'Review Findings
Range("O2") = thiswb.Sheets("TANF Computation").Range("B72") 'Benefit Amt
Range("P2") = thiswb.Sheets("TANF Computation").Range("C69") 'Benefit Determination
Range("Q2") = thisws.Range("AO10") 'Amt of Error

'reformat street address
    Range("Z2") = StrConv(thisws.Range("B3"), vbProperCase)
'Reformat city and state
    temparry = Split(thisws.Range("B4"), ",")
    TempStr = StrConv(Trim(temparry(LBound(temparry))), vbProperCase)

Range("AA2") = TempStr & ", " & Trim(temparry(UBound(temparry)))

End If
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Going out to website with list of Executive Directors
        FullTextFileName = datasourcews.Range("U6")
            
            'ActiveWorkbook.FollowHyperlink (FullTextFileName)
            Workbooks.Open fileName:=FullTextFileName, UpdateLinks:=False, ReadOnly:=True
            
            'Allegheny Districts are formatted as dates, change them to 02/XX format
            For irow = 1 To 100
                If IsDate(Range("A" & irow)) Then
                    TempStr = WorksheetFunction.Text(Range("A" & irow), "mm/d/yyyy")
                    Range("A" & irow).NumberFormat = "@"
                    Range("A" & irow) = Left(TempStr, 4)
                End If
            Next irow
            
            For irow = 1 To 1000
                If Range("A" & irow) = countylookup Then
                    datasourcews.Range("Y2") = Range("D" & irow)
                End If
            Next irow
            ActiveWorkbook.Close False
            
            datasourcewb.Close True
            
   
        findmemo = "TANF Pending Memo for Review Number " & review_number & " for Sample Month " & sample_month & ".doc"
    
    
    'Copy template finding memo to active directory
    SrceFile = PathStr & "Finding Memo\TANF Pending Letter.docx"
    tempfindmemoname = sPath & "\FM temp.docx"
    DestFile = tempfindmemoname
    FileCopy SrceFile, DestFile
    
    'Check to see if findings memo has already been created.
    If Dir(sPath & "\" & findmemo) <> "" Then
        Response = MsgBox("A TANF Pending memo has already been created. Do you want to overwrite this memo? No memo will be created if you click NO", Buttons:=vbYesNo)
    ' If statement to check if the NO button was selected.
      If Response = vbNo Then
        End
      End If
    End If

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

    .ActiveDocument.Content.Font.Name = "Arial"
    '.ActiveDocument.Content.Font.Size = 11

    .ActiveDocument.SaveAs sPath & "\" & findmemo
    .ActiveDocument.Close
    
    .Documents(tempfindmemoname).Close SaveChanges:=False

End With
appWD.Quit
Set appWD = Nothing

'delete temp files
'Kill tempfindmemoname
'Kill tempsourcename

Application.StatusBar = False
Application.ScreenUpdating = True
Application.Visible = True

MsgBox "A TANF Pending Memo has been saved as " & sPath & "\" & findmemo & _
". Please close the Internet Explorer with the CAO Directory. - Thank you!"

End Sub

Sub TANFSchool()

'Create Finding Memo
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

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'TANF
If review_type = 1 Then

' Get county number and name in correct format
countynum = Right(thisws.Range("U10"), 2)
districtnum = thisws.Range("Y10")
countylookup = countyformat(countynum, thisws.Range("O4"), districtnum)
Range("AK2") = thisws.Range("O4") & " Assistance Office"
'Look for counties with districts
'If countynum = "02" Or countynum = "23" Or countynum = "40" Or _
'        countynum = "51" Or countynum = "63" Or countynum = "65" Then
'    Range("AK2") = thisws.Range("O4") & " Office"
'Else
'    Range("AK2") = thisws.Range("O4") & " CAO"
'End If

'Client Information
Range("B2") = StrConv(thisws.Range("B2"), vbProperCase) 'Client Name
Range("C2") = Right(thisws.Range("U10"), 2) & "/" & thisws.Range("I10") 'County and Case Number
Range("D2") = thisws.Range("A10") 'Review Number
review_number = Range("D2")
Range("E2") = Left(thisws.Range("AB10"), 2) & "/" & Right(thisws.Range("AB10"), 4) 'Review Month
sample_month = thisws.Range("AB10")
Range("A7") = Val(thisws.Range("AO3") & thisws.Range("AP3")) 'Reviewer Number
Range("G2") = Range("C7") & "/" & thisws.Range("AG2") 'Supervisor and Reviewer Name
Range("E7") = review_type
Range("F7") = review_type
Range("R2") = thisws.Range("J61") & " - " & thisws.Range("O61") & " - " & thisws.Range("T61") 'Element, Nature, & Cause Code
Range("E17") = Val(thisws.Range("AL10")) 'Review Findings
Range("O2") = thiswb.Sheets("TANF Computation").Range("B72") 'Benefit Amt
Range("P2") = thiswb.Sheets("TANF Computation").Range("C69") 'Benefit Determination
Range("Q2") = thisws.Range("AO10") 'Amt of Error

'reformat street address
    Range("Z2") = StrConv(thisws.Range("B3"), vbProperCase)
'Reformat city and state
    temparry = Split(thisws.Range("B4"), ",")
    TempStr = StrConv(Trim(temparry(LBound(temparry))), vbProperCase)

Range("AA2") = TempStr & ", " & Trim(temparry(UBound(temparry)))

End If
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Going out to website with list of Executive Directors
        FullTextFileName = datasourcews.Range("U6")
            
            'ActiveWorkbook.FollowHyperlink (FullTextFileName)
            Workbooks.Open fileName:=FullTextFileName, UpdateLinks:=False, ReadOnly:=True
            
            'Allegheny Districts are formatted as dates, change them to 02/XX format
            For irow = 1 To 100
                If IsDate(Range("A" & irow)) Then
                    TempStr = WorksheetFunction.Text(Range("A" & irow), "mm/d/yyyy")
                    Range("A" & irow).NumberFormat = "@"
                    Range("A" & irow) = Left(TempStr, 4)
                End If
            Next irow
            
            For irow = 1 To 1000
                If Range("A" & irow) = countylookup Then
                    datasourcews.Range("Y2") = Range("D" & irow)
                End If
            Next irow
            ActiveWorkbook.Close False
            
            datasourcewb.Close True
            
   
        findmemo = "TANF School Verification Memo for Review Number " & review_number & " for Sample Month " & sample_month & ".doc"
    
    
    'Copy template finding memo to active directory
    SrceFile = PathStr & "Finding Memo\TANF School Verification Letter.docx"
    tempfindmemoname = sPath & "\FM temp.docx"
    DestFile = tempfindmemoname
    FileCopy SrceFile, DestFile
    
    'Check to see if findings memo has already been created.
    If Dir(sPath & "\" & findmemo) <> "" Then
        Response = MsgBox("A TANF School Verification Memo has already been created. Do you want to overwrite this memo? No memo will be created if you click NO", Buttons:=vbYesNo)
    ' If statement to check if the NO button was selected.
      If Response = vbNo Then
        End
      End If
    End If

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

    .ActiveDocument.Content.Font.Name = "Arial"
    '.ActiveDocument.Content.Font.Size = 11

    .ActiveDocument.SaveAs sPath & "\" & findmemo
    .ActiveDocument.Close
    
    .Documents(tempfindmemoname).Close SaveChanges:=False

End With
appWD.Quit
Set appWD = Nothing

'delete temp files
'Kill tempfindmemoname
'Kill tempsourcename

Application.StatusBar = False
Application.ScreenUpdating = True
Application.Visible = True

MsgBox "A TANF School Verification Memo has been saved as " & sPath & "\" & findmemo & _
". Please close the Internet Explorer with the CAO Directory. - Thank you!"

End Sub



