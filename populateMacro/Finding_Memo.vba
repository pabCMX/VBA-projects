Attribute VB_Name = "Finding_Memo"
Sub ShowSelectForms()
    SelectForms.Show
End Sub
Sub ShowMASelectForms()
    MASelectForms.Show
End Sub
Sub timeliness_switch()
Dim thisws As Worksheet, thiswb As Workbook

Set thisws = ActiveSheet
Set thiswb = ActiveWorkbook

'check if this is an info or finding memo
    If thisws.Range("K149").Value > 23 And thisws.Range("K149").Value < 31 And _
            thisws.Range("C149").Value = 2 Then
        Call Time("Info")
        Call Timeliness("Application")
    ElseIf thisws.Range("K149").Value > 23 And thisws.Range("K149").Value < 31 And _
            thisws.Range("C149").Value <> 2 Then
        Call Time("Info")
    ElseIf (thisws.Range("K149").Value > 10 And thisws.Range("K149").Value < 14) And _
            thisws.Range("C149").Value = 2 Then
        Call Timeliness("Renewal")
        Call Timeliness("Application")
    ElseIf (thisws.Range("K149").Value > 10 And thisws.Range("K149").Value < 14) And _
            thisws.Range("C149").Value <> 2 Then
        Call Timeliness("Renewal")
    ElseIf thisws.Range("C149").Value = 2 Or 3 Then
        Call Timeliness("Application")
    Else
        MsgBox "This case doesn't contain a finding or a client caused info action. Check #68 and #70 in Section 6."
    End If
    
End Sub

Sub Timeliness(finding_type As String)
'Create Timeliness Finding Memo
Dim appWD As Object
Dim review_number As Long
Dim sample_month As String, PathStr As String, findmemo As String
Dim countynum As String, districtnum As String
Dim datasourcewb As Workbook, datasourcews As Worksheet
Dim DSwb As Workbook, DSws As Worksheet

'Find path to Examiner's files on this PC
Set WshNetwork = CreateObject("WScript.Network")
Set oDrives = WshNetwork.EnumNetworkDrives

Dim DLetter As String, DUNC As String
Dim thisws As Worksheet, thiswb As Workbook
Dim sPath As String, tempsourcename As String

sPath = ActiveWorkbook.Path

Set thisws = ActiveSheet
Set thiswb = ActiveWorkbook
thiswb.SAVE

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

Application.StatusBar = "Starting processing ..."

SrceFile = PathStr & "Finding Memo\SNAP Timeliness QC Finding Memo.xlsm"

'Copy finding memo template to active directory
tempsourcename = sPath & "\FM Temp.xlsm"

DestFile = tempsourcename
FileCopy SrceFile, DestFile

'Open spreadsheet where all the data is stored to populate findings memo
Workbooks.Open fileName:=tempsourcename, UpdateLinks:=False

Set datasourcewb = ActiveWorkbook
Set datasourcews = ActiveWorkbook.Sheets("Findings Memo")

'Copy template data source to active directory
tempsourcenameDS = sPath & "\FM DS Temp.xlsx"
FileCopy PathStr & "Finding Memo\Finding Memo Data Source.xlsx", tempsourcenameDS

'Open spreadsheet where all the data is stored to populate findings memo
Workbooks.Open fileName:=tempsourcenameDS, UpdateLinks:=False

Set DSwb = ActiveWorkbook
Set DSws = ActiveSheet

'set file name to CAO List in Finding Memo Data Source file
FullTextFileName = DSws.Range("U6")

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////

'Today Date
    datasourcews.Range("J1") = Date
'Client Name
    datasourcews.Range("A5") = StrConv(thisws.Range("B4"), vbProperCase)
'reformat street address
    If thisws.Range("B6") <> "" Then
        datasourcews.Range("A6") = StrConv(thisws.Range("B5"), vbProperCase) & ", " & StrConv(thisws.Range("B6"), vbProperCase)
    Else
        datasourcews.Range("A6") = StrConv(thisws.Range("B5"), vbProperCase)
    End If
'Reformat city and state
    temparry = Split(thisws.Range("B7"), ",")
    TempStr = StrConv(Trim(temparry(LBound(temparry))), vbProperCase)
    datasourcews.Range("A7") = TempStr & ", " & Trim(temparry(UBound(temparry)))
'county/case number
    datasourcews.Range("D5") = thisws.Range("C155") & thisws.Range("D155") & "/" & Left(thisws.Range("I18"), 9)
'county number - hidden on sheet
    datasourcews.Range("D8") = thisws.Range("C155") & thisws.Range("D155")
'district number - hidden on sheet
    datasourcews.Range("E8") = thisws.Range("E155") & thisws.Range("F155")
'processing center
    datasourcews.Range("D7") = thisws.Range("X18") & thisws.Range("B153") & thisws.Range("C153")
'review number
    datasourcews.Range("F5") = thisws.Range("A18")
    review_number = datasourcews.Range("F5")
'review month
    datasourcews.Range("H5") = thisws.Range("AD18") & "/" & thisws.Range("AG18")
    sample_month = thisws.Range("AD18") & thisws.Range("AG18")
'County Name
    datasourcews.Range("J6") = thisws.Range("M5")
'Error Type
    If finding_type = "Application" Then
        datasourcews.Range("G14") = "Untimely Application SNAP"
        findmemo = "QC FINDING Application Timeliness Review Number " & review_number & " for Sample Month " & sample_month & ".xlsm"
    Else
        datasourcews.Range("G14") = "Untimely Renewal SNAP"
        findmemo = "QC FINDING Renewal Timeliness Review Number " & review_number & " for Sample Month " & sample_month & ".xlsm"
    End If
    
'Get county number and name in correct format
    countynum = Right(thisws.Range("X18"), 2)
    districtnum = thisws.Range("B153") & thisws.Range("C153")
    countylookup = countyformat(countynum, thisws.Range("M5"), districtnum)
    areanum = Left(thisws.Range("X18"), 1)
    
'use ds worksheet to find out qc manager, supervisor and program manager
review_type = 5
DSws.Range("A7") = Val(thisws.Range("AJ5") & thisws.Range("AK5"))
DSws.Range("G2") = DSws.Range("C7") & "/" & thisws.Range("AE5")
DSws.Range("F7") = review_type
DSws.Range("E7") = review_type

'put in county and district to get email address
    DSws.Range("A50") = datasourcews.Range("D8")
    If datasourcews.Range("E8") = "" Then
        DSws.Range("B50") = "BB"
    Else
        DSws.Range("B50") = WorksheetFunction.Text(datasourcews.Range("E8"), "00")
    End If
'cao email - hidden on sheet
    datasourcews.Range("F8") = DSws.Range("C50")
'program manager email - hidden on sheet
    datasourcews.Range("I11") = DSws.Range("N7")
    
'fill in the above information in finding memo
    datasourcews.Range("A10") = DSws.Range("H2") ' Program Manager
    datasourcews.Range("D10") = DSws.Range("AP2") ' Supervisor
    datasourcews.Range("I10") = DSws.Range("AM2") ' Examiner
    
'close without saving DS file and delete file
    DSwb.Close False
    Kill tempsourcenameDS
    
'Open CC list file
    Workbooks.Open fileName:=PathStr & "Finding Memo\Courtesy Copy for Memos.xlsx", UpdateLinks:=False, ReadOnly:=True
    Set ccwb = ActiveWorkbook
    Set ccws = ccwb.Worksheets(areanum)

'Form cc string
    ccstring = vbNewLine & vbNewLine & "cc: "
    ccstring = ccstring & ccws.Range("D4") & vbNewLine 'Area Manager
    ccstring = ccstring & ccws.Range("D5") & vbNewLine 'Area Staff Asst
    ccrow = 6
    'check area number - area 1 has QA person
    If areanum = "1" Then
        ccstring = ccstring & ccws.Range("D" & ccrow) & vbNewLine 'QA for Area 1
        ccrow = ccrow + 1
    End If

    ccstring = ccstring & datasourcews.Range("A10") & vbNewLine 'Office manager
    ccrow = ccrow + 1
    
    'column D is for SNAP Findings column in Courtesy Copy for Memo file
    cc_col = "D"
    
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
    
'Close cc workbook
    ccwb.Close False
' Add cc string to textbox
    datasourcews.Shapes("TextBox 1").TextFrame.Characters.Text = datasourcews.Shapes("TextBox 1").TextFrame.Characters.Text & ccstring
    
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Going out to website with list of Executive Directors
        'FullTextFileName = DSws.Range("U6")
            
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
                    datasourcews.Range("J5") = Range("D" & irow)
                    Exit For
                End If
            Next irow
            ActiveWorkbook.Close False
            
    'Check to see if findings memo has already been created.
    If Dir(sPath & "\" & findmemo) <> "" Then
        Response = MsgBox("A timeliness " & finding_type & " findings memo has already been created. Do you want to overwrite this memo? No memo will be created if you click NO", Buttons:=vbYesNo)
    ' If statement to check if the NO button was selected.
      If Response = vbNo Then
        End
      Else
      'if finding memo exists, then delete it and create new memo
        Kill sPath & "\" & findmemo
      End If
    End If
    
'save finding memo and then rename it

datasourcewb.Close True
Name tempsourcename As sPath & "\" & findmemo

Application.StatusBar = False
Application.ScreenUpdating = True
Application.Visible = True

MsgBox "A Timeliness " & finding_type & "  Finding Memo has been saved as " & sPath & "\" & findmemo & _
"."
End Sub

Sub Finding_Memo_sub()

'Create Finding Memo
Dim appWD As Object
Dim review_number As Long
Dim sample_month As String, PathStr As String, findmemo As String
Dim countynum As String, districtnum As String
Dim datasourcewb As Workbook, datasourcews As Worksheet
Dim DSwb As Workbook, DSws As Worksheet
Dim thiscompws As Worksheet, datasourcecompws As Worksheet

'Find path to Examiner's files on this PC
Set WshNetwork = CreateObject("WScript.Network")
Set oDrives = WshNetwork.EnumNetworkDrives

Dim DLetter As String, DUNC As String
Dim thisws As Worksheet, thiswb As Workbook
Dim sPath As String, tempsourcename As String

sPath = ActiveWorkbook.Path

Set thisws = ActiveSheet
Set thiswb = ActiveWorkbook
thiswb.SAVE

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

'determing what kind of review it is by the first number in the review
review_type = Left(thisws.Name, 1)

'SNAP Positive
If review_type = 5 Then
    'check if error case
    If thisws.Range("K22") = 1 Then
        MsgBox "Finding = 1. This case is not an error case."
        End
    End If
    SrceFile = PathStr & "Finding Memo\SNAP Positive QC Finding Memo.xlsm"
'SNAP Negative
ElseIf review_type = 6 Then
    'check if error case
    'If thisws.Range("AF29") = 1 Then 'code for old snap neg
    If thisws.Range("M29") = 1 Then 'code for new snap neg 11/2014
        MsgBox "Finding = 1. This case is not an error case."
        End
    End If
    SrceFile = PathStr & "Finding Memo\SNAP Negative QC Finding Memo.xlsm"
'TANF
ElseIf review_type = 1 Then
    'check if error case
    If thisws.Range("AL10") = 1 Then
        MsgBox "Finding = 1. This case is not an error case."
        End
    End If
    SrceFile = PathStr & "Finding Memo\TANF QC Finding Memo.xlsm"
'PE review
'ElseIf Left(thisws.Name, 2) = "24" Then
ElseIf review_type = 9 Then
    SrceFile = PathStr & "Finding Memo\GA QC Finding Memo.xlsm"
'MA
ElseIf review_type = 2 Then
    SrceFile = PathStr & "Finding Memo\MA Positive QC Finding Memo.xlsm"
'MA Negative
ElseIf review_type = 8 Then
    SrceFile = PathStr & "Finding Memo\MA Negative QC Finding Memo.xlsm"
End If

'Copy finding memo template to active directory
tempsourcename = sPath & "\FM Temp.xlsm"

DestFile = tempsourcename
FileCopy SrceFile, DestFile

'Open spreadsheet where all the data is stored to populate findings memo
Workbooks.Open fileName:=tempsourcename, UpdateLinks:=False

Set datasourcewb = ActiveWorkbook
Set datasourcews = ActiveWorkbook.Sheets("Findings Memo")

'Copy template data source to active directory
tempsourcenameDS = sPath & "\FM DS Temp.xlsx"
FileCopy PathStr & "Finding Memo\Finding Memo Data Source.xlsx", tempsourcenameDS

'Open spreadsheet where all the data is stored to populate findings memo
Workbooks.Open fileName:=tempsourcenameDS, UpdateLinks:=False

Set DSwb = ActiveWorkbook
Set DSws = ActiveSheet

'set file name to CAO List in Finding Memo Data Source file
FullTextFileName = DSws.Range("U6")

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////
'SNAP Positive
If review_type = 5 Then

'Today Date
    datasourcews.Range("J1") = Date
'Client Name
    datasourcews.Range("A5") = StrConv(thisws.Range("B4"), vbProperCase)
'reformat street address
    If thisws.Range("B6") <> "" Then
        datasourcews.Range("A6") = StrConv(thisws.Range("B5"), vbProperCase) & ", " & StrConv(thisws.Range("B6"), vbProperCase)
    Else
        datasourcews.Range("A6") = StrConv(thisws.Range("B5"), vbProperCase)
    End If
'Reformat city and state
    temparry = Split(thisws.Range("B7"), ",")
    TempStr = StrConv(Trim(temparry(LBound(temparry))), vbProperCase)
    datasourcews.Range("A7") = TempStr & ", " & Trim(temparry(UBound(temparry)))
'county/case number
    datasourcews.Range("D5") = thisws.Range("C155") & thisws.Range("D155") & "/" & Left(thisws.Range("I18"), 9)
'county number - hidden on sheet
    datasourcews.Range("D8") = thisws.Range("C155") & thisws.Range("D155")
'district number - hidden on sheet
    datasourcews.Range("E8") = thisws.Range("E155") & thisws.Range("F155")
'processing center
    datasourcews.Range("D7") = thisws.Range("X18") & thisws.Range("B153") & thisws.Range("C153")
'review number
    datasourcews.Range("F5") = thisws.Range("A18")
    review_number = datasourcews.Range("F5")
'review month
    datasourcews.Range("H5") = thisws.Range("AD18") & "/" & thisws.Range("AG18")
    sample_month = thisws.Range("AD18") & thisws.Range("AG18")
'County Name
    datasourcews.Range("J6") = thisws.Range("M5")
'Benefit Amount
    datasourcews.Range("B12") = thisws.Range("P22")
'Error Amount
    datasourcews.Range("I12") = thisws.Range("Y22")
'QC Determined Benefit Amount
    'datasourcews.Range("E12") = thiswb.Sheets("TANF Computation").Range("C71")
'Error Type
    If thisws.Range("K22") = 2 Then
        datasourcews.Range("L12") = "Overissuance"
    ElseIf thisws.Range("K22") = 3 Then
        datasourcews.Range("L12") = "Underissuance"
    ElseIf thisws.Range("K22") = 4 Then
        datasourcews.Range("L12") = "Ineligible"
    End If
    
'use ds worksheet to find out qc manager, supervisor and program manager
DSws.Range("A7") = Val(thisws.Range("AJ5") & thisws.Range("AK5")) 'examiner id
reviewerid = DSws.Range("A7")
DSws.Range("G2") = DSws.Range("C7") & "/" & thisws.Range("AE5") 'supervisor/examiner name
DSws.Range("F7") = review_type
DSws.Range("E7") = review_type

'put in county and district to get email address
    DSws.Range("A50") = datasourcews.Range("D8")
    If datasourcews.Range("E8") = "" Then
        DSws.Range("B50") = "BB"
    Else
        DSws.Range("B50") = WorksheetFunction.Text(datasourcews.Range("E8"), "00")
    End If
    
'cao email - hidden on sheet
    datasourcews.Range("F8") = DSws.Range("C50")
'program manager email - hidden on sheet
    datasourcews.Range("I11") = DSws.Range("N7")
    
'fill in the above information in finding memo
    datasourcews.Range("A10") = DSws.Range("H2")  ' Program Manager
    datasourcews.Range("D10") = DSws.Range("AP2") ' Supervisor
    datasourcews.Range("I10") = DSws.Range("AM2") ' Examiner
    
'close without saving DS file and then delete it
    DSwb.Close False
    Kill tempsourcenameDS
    
'Loop thru Section 2 to get the element, nature and cause codes
    enc_codes = ""
    For irow = 29 To 43 Step 2
        If thisws.Range("B" & irow) = "" Then
            Exit For
        Else
            If enc_codes <> "" Then
                enc_codes = enc_codes & vbNewLine
            End If
            enc_codes = enc_codes & thisws.Range("B" & irow) & " - " & thisws.Range("G" & irow) & " - " & thisws.Range("K" & irow)
        End If
    Next irow

'Put element, nature and cause codes on finding memo
    datasourcews.Range("G16") = enc_codes

'Look up in FS Workbook for all buttons labeled 2 or 3 which are selected
    txtbox_texts = find_texts_fs(thiswb.Sheets("FS Workbook"))
'Add text box with labels found in FS Workbook
    Set shpTemp = datasourcews.Shapes.AddTextbox(msoTextOrientationHorizontal, _
            5, 250, 500, 500)
            
'set text box name and add text
    With shpTemp
        .Name = "TBox1"
        .TextFrame2.AutoSize = msoAutoSizeShapeToFitText
        'add text to text box
        .TextFrame.Characters.Text = txtbox_texts
    End With

' Get county number and name in correct format
    countynum = Right(thisws.Range("X18"), 2)
    districtnum = thisws.Range("B153") & thisws.Range("C153")
    countylookup = countyformat(countynum, thisws.Range("M5"), districtnum)
    areanum = Left(thisws.Range("X18"), 1)
    
'Open CC list file
    Workbooks.Open fileName:=PathStr & "Finding Memo\Courtesy Copy for Memos.xlsx", UpdateLinks:=False, ReadOnly:=True
    Set ccwb = ActiveWorkbook
    Set ccws = ccwb.Worksheets(areanum)

'Form cc string
    ccstring = vbNewLine & vbNewLine & "cc: "
    ccstring = ccstring & ccws.Range("D4") & vbNewLine 'Area Manager
    ccstring = ccstring & ccws.Range("D5") & vbNewLine 'Area Staff Asst
    ccrow = 6
    'check area number
    If areanum = "1" Then
        ccstring = ccstring & ccws.Range("D" & ccrow) & vbNewLine 'QA for Area 1
        ccrow = ccrow + 1
    End If
    
    ccstring = ccstring & datasourcews.Range("A10") & vbNewLine 'Office manager
    ccrow = ccrow + 1
    
    'Column D is SNAP Findings column in Courtesy Copy for Memos file
    cc_col = "D"
    
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
    
'Close cc workbook
    ccwb.Close False
' Add cc string to textbox
    datasourcews.Shapes("TBox1").TextFrame.Characters.Text = datasourcews.Shapes("TBox1").TextFrame.Characters.Text & ccstring

'Format top line as bold underline
    datasourcews.Shapes.Range(Array("TBox1")).Select
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 24).Font.Bold = _
                msoTrue
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 24).Font. _
                UnderlineStyle = msoUnderlineSingleLine
                 
'Copy computation sheet to findings memo
'check if comp sheet exists
    On Error Resume Next
    Set thiscompws = thiswb.Sheets("FS Computation")
    On Error GoTo 0
    If thiscompws Is Nothing Then
        'if comp sheet doesn't exists alert user
        MsgBox "FS Computation sheet doesn't exist in file."
    Else
        'create new tab and name for fs comp sheet
        datasourcewb.Sheets.Add After:=Sheets(Sheets.Count)
        Set datasourcecompws = ActiveSheet
        datasourcecompws.Name = "FS Computation"
        'copy visible comp sheet cells
        thiscompws.Range("A1:AK71").Copy
        'select findings memo fs comp sheet tab
        datasourcecompws.Select
        'paste image of fs comp sheet
        ActiveSheet.Pictures.Paste.Select
        Application.CutCopyMode = False
        'set fill color of image to white
        ActiveSheet.Shapes.Range(Array("Picture 1")).Select
        With Selection.ShapeRange.Fill
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorBackground1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = 0
            .Transparency = 0
            .Solid
        End With
        'select the findings memo tab
        datasourcews.Select
    End If
    

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'SNAP Negative
ElseIf review_type = 6 Then

'Today Date
    datasourcews.Range("J1") = Date
'Client Name
    datasourcews.Range("A5") = StrConv(thisws.Range("C8"), vbProperCase)
'reformat street address
        datasourcews.Range("A6") = StrConv(thisws.Range("C12"), vbProperCase)
'Reformat city and state
    temparry = Split(thisws.Range("C13"), ",")
    TempStr = StrConv(Trim(temparry(LBound(temparry))), vbProperCase)
    datasourcews.Range("A7") = TempStr & ", " & Trim(temparry(UBound(temparry)))
'county/case number
  If thisws.Range("C20") = "06" Then
    datasourcews.Range("D5") = thisws.Range("F56") & thisws.Range("G56") & "/" & Left(thisws.Range("L20"), 9)
  Else
    datasourcews.Range("D5") = thisws.Range("F56") & thisws.Range("G56") & "/" & Right(thisws.Range("L20"), 8)
  End If
'county number
    datasourcews.Range("D8") = Right(thisws.Range("AA20"), 2)
'district number
    datasourcews.Range("E8") = thisws.Range("AD20") & thisws.Range("AE20")
    'datasourcews.Range("E8") = thisws.Range("H56") & thisws.Range("I56")
'processing center
    'datasourcews.Range("D7") = thisws.Range("E50") & thisws.Range("F50") & thisws.Range("G50") & thisws.Range("H50") & thisws.Range("I50") 'code for old snap neg
    'datasourcews.Range("D7") = thisws.Range("E56") & thisws.Range("F56") & thisws.Range("G56") & thisws.Range("H56") & thisws.Range("I56") 'code for new snap neg - 11/2014
    datasourcews.Range("D7") = thisws.Range("AA20") & thisws.Range("C56") & thisws.Range("D56")  'code for new snap neg - 11/2022
'review number
    datasourcews.Range("F5") = thisws.Range("C20")
    review_number = datasourcews.Range("F5")
'review month
    datasourcews.Range("H5") = thisws.Range("AF20") & "/" & thisws.Range("AI20")
    sample_month = thisws.Range("AF20") & thisws.Range("AI20")
'County Name
    datasourcews.Range("J6") = thisws.Range("C1")
'Error Type
    If thisws.Range("AE24") = 1 Then
        datasourcews.Range("F12") = "Invalid Denial"
    ElseIf thisws.Range("AE24") = 2 Then
        datasourcews.Range("F12") = "Invalid Termination"
    ElseIf thisws.Range("AE24") = 3 Then
        datasourcews.Range("F12") = "Invalid Suspension"
    End If
    
'get first Element Code and Nature Code
    'datasourcews.Range("G16") = thisws.Range("E41") & " - " & thisws.Range("W41") 'old code
    datasourcews.Range("G16") = thisws.Range("E47") & " - " & thisws.Range("W47") 'new code - 11/2014
    
'use ds worksheet to find out qc manager, supervisor and program manager
    DSws.Range("A7") = Val(thisws.Range("W17")) 'examiner id
    reviewerid = DSws.Range("A7")
    DSws.Range("G2") = DSws.Range("C7") & "/" & thisws.Range("Q17")
    DSws.Range("F7") = review_type
    DSws.Range("E7") = review_type

'put in county and district to get email address
    DSws.Range("A50") = datasourcews.Range("D8")
    If datasourcews.Range("E8") = "" Then
        DSws.Range("B50") = "BB"
    Else
        DSws.Range("B50") = WorksheetFunction.Text(datasourcews.Range("E8"), "00")
    End If
    
'cao email - hidden on sheet
    datasourcews.Range("F8") = DSws.Range("C50")
'program manager email - hidden on sheet
    datasourcews.Range("I11") = DSws.Range("N7")
    
'fill in the above information in finding memo
    datasourcews.Range("A10") = DSws.Range("H2")  ' Program Manager
    datasourcews.Range("D10") = DSws.Range("AP2") ' Supervisor
    datasourcews.Range("I10") = DSws.Range("AM2") ' Examiner
    
'close without saving DS file
    DSwb.Close False
    Kill tempsourcenameDS

' Get county number and name in correct format
    countynum = thisws.Range("F56") & thisws.Range("G56")
    districtnum = thisws.Range("H56") & thisws.Range("I56")
    countylookup = countyformat(countynum, thisws.Range("C1"), districtnum)
    areanum = thisws.Range("E56")

'Open CC list file
    Workbooks.Open fileName:=PathStr & "Finding Memo\Courtesy Copy for Memos.xlsx", UpdateLinks:=False, ReadOnly:=True
    Set ccwb = ActiveWorkbook
    Set ccws = ccwb.Worksheets(areanum)

'Form cc string
    ccstring = vbNewLine & vbNewLine & "cc: "
    ccstring = ccstring & ccws.Range("D4") & vbNewLine 'Area Manager
    ccstring = ccstring & ccws.Range("D5") & vbNewLine 'Area Staff Asst
    ccrow = 6
    'check area number
    If areanum = "1" Then
        ccstring = ccstring & ccws.Range("D" & ccrow) & vbNewLine 'QA for Area 1
        ccrow = ccrow + 1
    End If
    
    ccstring = ccstring & datasourcews.Range("A10") & vbNewLine 'Office manager
    ccrow = ccrow + 1
    
    'Column D is SNAP Findings column in Courtesy Copy for Memos file
    cc_col = "D"
    
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
    
'Close cc workbook
    ccwb.Close False
' Add cc string to textbox
    datasourcews.Shapes("TextBox 1").TextFrame.Characters.Text = datasourcews.Shapes("TextBox 1").TextFrame.Characters.Text & ccstring

'Format top line as bold underline
    datasourcews.Shapes.Range(Array("TextBox 1")).Select
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 24).Font.Bold = _
                msoTrue
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 24).Font. _
                UnderlineStyle = msoUnderlineSingleLine

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'TANF
ElseIf review_type = 1 Then

'Today Date
    datasourcews.Range("J1") = Date
'Client Name
    datasourcews.Range("A5") = StrConv(thisws.Range("B2"), vbProperCase)
'reformat street address
    'If thisws.Range("B5") <> "" Then
    '    datasourcews.Range("A6") = StrConv(thisws.Range("B3"), vbProperCase) & ", " & StrConv(thisws.Range("B4"), vbProperCase)
    'Else
        datasourcews.Range("A6") = StrConv(thisws.Range("B3"), vbProperCase)
    'End If
'Reformat city and state
    temparry = Split(thisws.Range("B4"), ",")
    TempStr = StrConv(Trim(temparry(LBound(temparry))), vbProperCase)
    datasourcews.Range("A7") = TempStr & ", " & Trim(temparry(UBound(temparry)))
'county/case number
    datasourcews.Range("D5") = Right(thisws.Range("U10"), 2) & "/" & thisws.Range("I10")
'county number - hidden on sheet
    datasourcews.Range("D8") = Right(thisws.Range("U10"), 2)
'district number - hidden on sheet
    datasourcews.Range("E8") = thisws.Range("Y10")
'processing center
    datasourcews.Range("D7") = thisws.Range("O85") & thisws.Range("P85") & thisws.Range("Q85") & thisws.Range("R85") & thisws.Range("S85")
'review number
    datasourcews.Range("F5") = thisws.Range("A10")
    review_number = datasourcews.Range("F5")
'review month
    datasourcews.Range("H5") = Left(thisws.Range("AB10"), 2) & "/" & Right(thisws.Range("AB10"), 4)
    sample_month = thisws.Range("AB10")
'County Name
    datasourcews.Range("J6") = thisws.Range("O4")
'Benefit Amount
    datasourcews.Range("B12") = thisws.Range("I20")
'Error Amount
    datasourcews.Range("I12") = thisws.Range("AO10")
'QC Determined Benefit Amount
    'datasourcews.Range("E12") = thisws.Range("C20") ***Jeannette asked this to be blank on 11.29.2021***
'Error Type
    If thisws.Range("AL10") = 2 Then
        datasourcews.Range("L12") = "Overissuance"
    ElseIf thisws.Range("AL10") = 3 Then
        datasourcews.Range("L12") = "Underissuance"
    ElseIf thisws.Range("AL10") = 4 Then
        datasourcews.Range("L12") = "Ineligible"
    End If
    
'use ds worksheet to find out qc manager, supervisor and program manager
    DSws.Range("A7") = Val(thisws.Range("AO3") & thisws.Range("AP3")) 'examiner id
    reviewerid = DSws.Range("A7")
    DSws.Range("G2") = DSws.Range("C7") & "/" & thisws.Range("AG2") 'supervisor/examiner name
    DSws.Range("F7") = review_type
    DSws.Range("E7") = review_type

'put in county and district to get email address
    DSws.Range("A50") = datasourcews.Range("D8")
    If datasourcews.Range("E8") = "" Then
        DSws.Range("B50") = "BB"
    Else
        DSws.Range("B50") = WorksheetFunction.Text(datasourcews.Range("E8"), "00")
    End If
    
'cao email - hidden on sheet
    datasourcews.Range("F8") = DSws.Range("C50")
'program manager email - hidden on sheet
    datasourcews.Range("I11") = DSws.Range("N7")
    
'fill in the above information in finding memo
    datasourcews.Range("A10") = DSws.Range("H2")  ' Program Manager
    datasourcews.Range("D10") = DSws.Range("AP2") ' Supervisor
    datasourcews.Range("I10") = DSws.Range("AM2") ' Examiner
    
'close without saving DS file and then delete it
    DSwb.Close False
    Kill tempsourcenameDS
    
'Loop thru Section 6 to get the element, nature and cause codes
    enc_codes = ""
    For irow = 61 To 67 Step 2
        If thisws.Range("F" & irow) = "" Then
            Exit For
        Else
            If enc_codes <> "" Then
                enc_codes = enc_codes & vbNewLine
            End If
            enc_codes = enc_codes & thisws.Range("J" & irow) & " - " & thisws.Range("O" & irow) & " - " & thisws.Range("T" & irow)
        End If
    Next irow

'Put element, nature and cause codes on finding memo
    datasourcews.Range("G16") = enc_codes

'Look up in TANF Workbook for all buttons labeled 2 or 3 which are selected
    txtbox_texts = find_texts_tanf(thiswb.Sheets("TANF Workbook"))

' Get county number and name in correct format
    countynum = Right(thisws.Range("U10"), 2)
    districtnum = thisws.Range("Y10")
    countylookup = countyformat(countynum, thisws.Range("O4"), districtnum)
    areanum = Left(thisws.Range("U10"), 1)

'Open CC list file
    Workbooks.Open fileName:=PathStr & "Finding Memo\Courtesy Copy for Memos.xlsx", UpdateLinks:=False, ReadOnly:=True
    Set ccwb = ActiveWorkbook
    Set ccws = ccwb.Worksheets(areanum)

'Form cc string
    ccstring = vbNewLine & vbNewLine & "cc: "
    ccstring = ccstring & ccws.Range("F4") & vbNewLine 'Area Manager
    ccstring = ccstring & ccws.Range("F5") & vbNewLine 'Area Staff Asst
    ccrow = 6
    'check area number
    If areanum = "1" Then
        ccstring = ccstring & ccws.Range("F" & ccrow) & vbNewLine 'QA for Area 1
        ccrow = ccrow + 1
    End If
    
    ccstring = ccstring & datasourcews.Range("A10") & vbNewLine 'Office manager
    ccrow = ccrow + 1
    
    'Column F is TANF Findings column in Courtesy Copy for Memos file
    cc_col = "F"
    
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
    
'Close cc workbook
    ccwb.Close False
' Add cc string to textbox
    datasourcews.Shapes("TextBox 1").TextFrame.Characters.Text = datasourcews.Shapes("TextBox 1").TextFrame.Characters.Text & txtbox_texts & ccstring

'Format top line as bold underline
    datasourcews.Shapes.Range(Array("TextBox 1")).Select
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 24).Font.Bold = _
                msoTrue
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 24).Font. _
                UnderlineStyle = msoUnderlineSingleLine
                
'Copy computation sheet to findings memo
'check if comp sheet exists
    On Error Resume Next
    Set thiscompws = thiswb.Sheets("TANF Computation")
    On Error GoTo 0
    If thiscompws Is Nothing Then
        'if comp sheet doesn't exists alert user
        MsgBox "TANF Computation sheet doesn't exist in file."
    Else
        'create new tab and name for tanf comp sheet
        datasourcewb.Sheets.Add After:=Sheets(Sheets.Count)
        Set datasourcecompws = ActiveSheet
        datasourcecompws.Name = "TANF Computation"
        'copy visible comp sheet cells
        thiscompws.Range("A1:N84").Copy
        'select findings memo fs comp sheet tab
        datasourcecompws.Select
        'paste image of fs comp sheet
        ActiveSheet.Pictures.Paste.Select
        Application.CutCopyMode = False
        'set fill color of image to white
        ActiveSheet.Shapes.Range(Array("Picture 1")).Select
        With Selection.ShapeRange.Fill
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorBackground1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = 0
            .Transparency = 0
            .Solid
        End With
        'select the findings memo tab
        datasourcews.Select
    End If
    
'GA
ElseIf review_type = 9 Then

'Today Date
    datasourcews.Range("J1") = Date
'Client Name
    datasourcews.Range("A5") = StrConv(thisws.Range("B2"), vbProperCase)
'reformat street address
    'If thisws.Range("B5") <> "" Then
    '    datasourcews.Range("A6") = StrConv(thisws.Range("B3"), vbProperCase) & ", " & StrConv(thisws.Range("B4"), vbProperCase)
    'Else
        datasourcews.Range("A6") = StrConv(thisws.Range("B3"), vbProperCase)
    'End If
'Reformat city and state
    temparry = Split(thisws.Range("B5"), ",")
    TempStr = StrConv(Trim(temparry(LBound(temparry))), vbProperCase)
    datasourcews.Range("A7") = TempStr & ", " & Trim(temparry(UBound(temparry)))
'county/case number
    datasourcews.Range("D5") = Right(thisws.Range("U10"), 2) & "/" & thisws.Range("I10")
'county number - hidden on sheet
    datasourcews.Range("D8") = Right(thisws.Range("U10"), 2)
'district number - hidden on sheet
    datasourcews.Range("E8") = thisws.Range("Y10")
'processing center
    datasourcews.Range("D7") = thisws.Range("O85") & thisws.Range("P85") & thisws.Range("Q85") & thisws.Range("R85") & thisws.Range("S85")
'review number
    datasourcews.Range("F5") = thisws.Range("A10")
    review_number = datasourcews.Range("F5")
'review month
    datasourcews.Range("H5") = Left(thisws.Range("AB10"), 2) & "/" & Right(thisws.Range("AB10"), 4)
    sample_month = thisws.Range("AB10")
'County Name
    datasourcews.Range("J6") = thisws.Range("O4")
'Benefit Amount
    datasourcews.Range("B12") = thisws.Range("I20")
'Error Amount
    datasourcews.Range("I12") = thisws.Range("AO10")
'QC Determined Benefit Amount
    datasourcews.Range("E12") = thisws.Range("C20")
'Error Type
    If thisws.Range("AL10") = 2 Then
        datasourcews.Range("L12") = "Overissuance"
    ElseIf thisws.Range("AL10") = 3 Then
        datasourcews.Range("L12") = "Underissuance"
    ElseIf thisws.Range("AL10") = 4 Then
        datasourcews.Range("L12") = "Ineligible"
    End If
    
'use ds worksheet to find out qc manager, supervisor and program manager
    DSws.Range("A7") = Val(thisws.Range("AO3") & thisws.Range("AP3")) 'examiner id
    reviewerid = DSws.Range("A7")
    DSws.Range("G2") = DSws.Range("C7") & "/" & thisws.Range("AG2") 'supervisor/examiner name
    DSws.Range("F7") = review_type
    DSws.Range("E7") = review_type

'put in county and district to get email address
    DSws.Range("A50") = datasourcews.Range("D8")
    If datasourcews.Range("E8") = "" Then
        DSws.Range("B50") = "BB"
    Else
        DSws.Range("B50") = WorksheetFunction.Text(datasourcews.Range("E8"), "00")
    End If
    
'cao email - hidden on sheet
    datasourcews.Range("F8") = DSws.Range("C50")
'program manager email - hidden on sheet
    datasourcews.Range("I11") = DSws.Range("N7")
    
'fill in the above information in finding memo
    datasourcews.Range("A10") = DSws.Range("F2")
    datasourcews.Range("D10") = DSws.Range("G2")
    datasourcews.Range("I10") = DSws.Range("H2")
    
'close without saving DS file and then delete it
    DSwb.Close False
    Kill tempsourcenameDS
    
'Loop thru Section 6 to get the element, nature and cause codes
    enc_codes = ""
    For irow = 61 To 67 Step 2
        If thisws.Range("F" & irow) = "" Then
            Exit For
        Else
            If enc_codes <> "" Then
                enc_codes = enc_codes & vbNewLine
            End If
            enc_codes = enc_codes & thisws.Range("J" & irow) & " - " & thisws.Range("O" & irow) & " - " & thisws.Range("T" & irow)
        End If
    Next irow

'Put element, nature and cause codes on finding memo
    datasourcews.Range("G16") = enc_codes

'Look up in TANF Workbook for all buttons labeled 2 or 3 which are selected
   ' txtbox_texts = find_texts_tanf(thiswb.Sheets("TANF Workbook"))

' Get county number and name in correct format
    countynum = Right(thisws.Range("U10"), 2)
    districtnum = thisws.Range("Y10")
    countylookup = countyformat(countynum, thisws.Range("O4"), districtnum)
    areanum = Left(thisws.Range("U10"), 1)

'Open CC list file
    Workbooks.Open fileName:=PathStr & "Finding Memo\Courtesy Copy for Memos.xlsx", UpdateLinks:=False, ReadOnly:=True
    Set ccwb = ActiveWorkbook
    Set ccws = ccwb.Worksheets(areanum)

'Form cc string
    ccstring = vbNewLine & vbNewLine & "cc: "
    ccstring = ccstring & ccws.Range("F4") & vbNewLine 'Area Manager
    ccstring = ccstring & ccws.Range("F5") & vbNewLine 'Area Staff Asst
    ccrow = 6
    'check area number
    If areanum = "1" Then
        ccstring = ccstring & ccws.Range("F" & ccrow) & vbNewLine 'QA for Area 1
        ccrow = ccrow + 1
    End If
    
    ccstring = ccstring & datasourcews.Range("A10") & vbNewLine 'Office manager
    ccrow = ccrow + 1
    
    'Column F is TANF Findings column in Courtesy Copy for Memos file
    cc_col = "F"
    
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
    
'Close cc workbook
    ccwb.Close False
' Add cc string to textbox
    datasourcews.Shapes("TextBox 1").TextFrame.Characters.Text = datasourcews.Shapes("TextBox 1").TextFrame.Characters.Text & txtbox_texts & ccstring

'Format top line as bold underline
    datasourcews.Shapes.Range(Array("TextBox 1")).Select
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 24).Font.Bold = _
                msoTrue
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 24).Font. _
                UnderlineStyle = msoUnderlineSingleLine
                
'Copy computation sheet to findings memo
'check if comp sheet exists
    On Error Resume Next
    Set thiscompws = thiswb.Sheets("GA Computation")
    On Error GoTo 0
    If thiscompws Is Nothing Then
        'if comp sheet doesn't exists alert user
        MsgBox "GA Computation sheet doesn't exist in file."
    Else
        'create new tab and name for tanf comp sheet
        datasourcewb.Sheets.Add After:=Sheets(Sheets.Count)
        Set datasourcecompws = ActiveSheet
        datasourcecompws.Name = "GA Computation"
        'copy visible comp sheet cells
        thiscompws.Range("A1:N84").Copy
        'select findings memo fs comp sheet tab
        datasourcecompws.Select
        'paste image of fs comp sheet
        ActiveSheet.Pictures.Paste.Select
        Application.CutCopyMode = False
        'set fill color of image to white
        ActiveSheet.Shapes.Range(Array("Picture 1")).Select
        With Selection.ShapeRange.Fill
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorBackground1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = 0
            .Transparency = 0
            .Solid
        End With
        'select the findings memo tab
        datasourcews.Select
    End If
    
End If
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Going out to website with list of Executive Directors
        'FullTextFileName = DSws.Range("U6")
            
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
                    datasourcews.Range("J5") = Range("D" & irow)
                    Exit For
                End If
            Next irow
            ActiveWorkbook.Close False
            
    If review_type = 2 Then
        sample_month = Left(sample_month, 2) & Right(sample_month, 4)
        findmemo = "QC FINDING Review Number " & review_number & " for Sample Month " & sample_month & ".xlsm"
    Else
        findmemo = "QC FINDING Review Number " & review_number & " for Sample Month " & sample_month & ".xlsm"
    End If
        
    'Check to see if findings memo has already been created.
    If Dir(sPath & "\" & findmemo) <> "" Then
        Response = MsgBox("A findings memo has already been created. Do you want to overwrite this memo? No memo will be created if you click NO", Buttons:=vbYesNo)
    ' If statement to check if the NO button was selected.
      If Response = vbNo Then
        End
      Else
      'if finding memo exists, then delete it and create new memo
        Kill sPath & "\" & findmemo
      End If
    End If
    
'save finding memo and then rename it

datasourcewb.Close True
Name tempsourcename As sPath & "\" & findmemo

Application.StatusBar = False
Application.ScreenUpdating = True
Application.Visible = True

MsgBox "A Finding Memo has been saved as " & sPath & "\" & findmemo & "."

End Sub

Sub MA_Finding_Memo_sub(memo_type As String)

'Create Finding Memo
Dim appWD As Object
Dim review_number As Long
Dim sample_month As String, PathStr As String, findmemo As String
Dim countynum As String, districtnum As String
Dim datasourcewb As Workbook, datasourcews As Worksheet
Dim DSwb As Workbook, DSws As Worksheet
Dim thiscompws As Worksheet, datasourcecompws As Worksheet

'Find path to Examiner's files on this PC
Set WshNetwork = CreateObject("WScript.Network")
Set oDrives = WshNetwork.EnumNetworkDrives

Dim DLetter As String, DUNC As String
Dim thisws As Worksheet, thiswb As Workbook
Dim sPath As String, tempsourcename As String

sPath = ActiveWorkbook.Path

Set thisws = ActiveSheet
Set thiswb = ActiveWorkbook
thiswb.SAVE

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

'determing what kind of review it is by the first number in the review
review_type = Left(thisws.Name, 1)

'what type of MA memo
Select Case memo_type
'MA Positive Finding
    Case "MA_Pos_Find"
        SrceFile = PathStr & "Finding Memo\MA Positive QC Finding Memo.xlsm"
        final_title = "MA Positive QC Finding Memo "
'MA Positive Deficiency
    Case "MA_Pos_Def"
        SrceFile = PathStr & "Finding Memo\MA Positive QC Deficiency Memo.xlsm"
        final_title = "MA Positive QC Deficiency Memo "
'MA Positive Information
    Case "MA_Pos_Info"
        SrceFile = PathStr & "Finding Memo\MA Positive QC Information Memo.xlsm"
        final_title = "MA Positive QC Information Memo "
'MA Negative Finding
    Case "MA_Neg_Find"
        SrceFile = PathStr & "Finding Memo\MA Negative QC Finding Memo.xlsm"
        final_title = "MA Negative QC Finding Memo "
'MA Negative Deficiency
    Case "MA_Neg_Def"
        SrceFile = PathStr & "Finding Memo\MA Negative QC Deficiency Memo.xlsm"
        final_title = "MA Negative QC Deficiency Memo "
'MA Negative Information
    Case "MA_Neg_Info"
        SrceFile = PathStr & "Finding Memo\MA Negative QC Information Memo.xlsm"
        final_title = "MA Negative QC Information Memo "
'MA PE Finding
    Case "MA_PE_Find"
        SrceFile = PathStr & "Finding Memo\MA PE QC Finding Memo.xlsm"
        final_title = "MA PE QC Finding Memo "
'MA PE Information
    Case "MA_PE_Info"
        SrceFile = PathStr & "Finding Memo\MA PE QC Information Memo.xlsm"
        final_title = "MA PE QC Information Memo "
'MA Positive Save Deficiency
Case "MA_Pos_SAVE"
        SrceFile = PathStr & "Finding Memo\MA Positive SAVE QC Deficiency Memo.xlsm"
        final_title = "MA Positive QC Finding Memo "
End Select

'temporary name for memo template
tempsourcename = sPath & "\FM Temp.xlsm"

'Copy finding memo template to active directory
DestFile = tempsourcename
FileCopy SrceFile, DestFile

'Open memo template spreadsheet where all the data is stored to populate memo
Workbooks.Open fileName:=tempsourcename, UpdateLinks:=False

'set shortcuts to memo template spreadsheet
Set datasourcewb = ActiveWorkbook
Set datasourcews = ActiveWorkbook.Sheets("Findings Memo")

'Copy DS template data source to active directory. Use this to find area manager, supervisor, etc.
tempsourcenameDS = sPath & "\FM DS Temp.xlsx"
FileCopy PathStr & "Finding Memo\Finding Memo Data Source.xlsx", tempsourcenameDS

'Open DS spreadsheet
Workbooks.Open fileName:=tempsourcenameDS, UpdateLinks:=False

'set shortcut to DS spreadsheet
Set DSwb = ActiveWorkbook
Set DSws = ActiveSheet

'set file name to CAO List in Finding Memo Data Source file
FullTextFileName = DSws.Range("U6")

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////
'For different reviews, populate upper part
Select Case Left(memo_type, 6)
    'Positive reviews
    Case "MA_Pos"
    'Today Date
        datasourcews.Range("J1") = Date
    'Client Name
        datasourcews.Range("A5") = StrConv(thisws.Range("B2"), vbProperCase)
    'reformat street address
        If thisws.Range("B4") <> "" Then
            datasourcews.Range("A6") = StrConv(thisws.Range("B3"), vbProperCase) & ", " & StrConv(thisws.Range("B4"), vbProperCase)
        Else
            datasourcews.Range("A6") = StrConv(thisws.Range("B3"), vbProperCase)
        End If
    'Reformat city and state
        temparry = Split(thisws.Range("B5"), ",")
        TempStr = StrConv(Trim(temparry(LBound(temparry))), vbProperCase)
        datasourcews.Range("A7") = TempStr & ", " & Trim(temparry(UBound(temparry)))
    'county/case number
        datasourcews.Range("D5") = Right(thisws.Range("AC10"), 2) & "/" & Left(thisws.Range("H10"), 9)
    'county number - hidden on sheet
        datasourcews.Range("D8") = Right(thisws.Range("AC10"), 2)
    'district number - hidden on sheet
        datasourcews.Range("E8") = thisws.Range("AH10")
    'review number
        datasourcews.Range("F5") = thisws.Range("A10") & thisws.Range("B10") & thisws.Range("C10") & thisws.Range("D10") & thisws.Range("E10") & thisws.Range("F10")
        review_number = datasourcews.Range("F5")
    'review month
        datasourcews.Range("H5") = thisws.Range("AM10")
        sample_month = Left(thisws.Range("AM10"), 2) & Right(thisws.Range("AM10"), 4)
    'County Name
        datasourcews.Range("J6") = thisws.Range("O4")
    'processing center
        datasourcews.Range("D7") = thisws.Range("W122") & thisws.Range("X122") & thisws.Range("Y122")
    
    'use ds worksheet to find out qc manager, supervisor and program manager
        DSws.Range("A7") = Val(thisws.Range("AO3") & thisws.Range("AP3")) 'examiner id
        reviewerid = DSws.Range("A7")
        DSws.Range("G2") = DSws.Range("C7") & "/" & DSws.Range("AM2") 'supervisor/examiner name
        DSws.Range("F7") = review_type
        DSws.Range("E7") = review_type

'put in county and district to get email address
    DSws.Range("A50") = datasourcews.Range("D8")
    If datasourcews.Range("E8") = "" Then
        DSws.Range("B50") = "BB"
    Else
        DSws.Range("B50") = WorksheetFunction.Text(datasourcews.Range("E8"), "00")
    End If
'cao email - hidden on sheet
    datasourcews.Range("F8") = DSws.Range("C50")
'program manager email - hidden on sheet
    datasourcews.Range("I11") = DSws.Range("N7")
    
    'fill in the above information in finding memo
    datasourcews.Range("A10") = DSws.Range("H2")  ' Program Manager
    datasourcews.Range("D10") = DSws.Range("AP2") ' Supervisor
    datasourcews.Range("I10") = DSws.Range("AM2") ' Examiner
    
    'close without saving DS file and then delete it
        DSwb.Close False
        Kill tempsourcenameDS
    
    ' Get county number and name in correct format
        countynum = Right(thisws.Range("AC10"), 2)
        districtnum = thisws.Range("AH10")
        countylookup = countyformat(countynum, thisws.Range("O4"), districtnum)
        areanum = Left(thisws.Range("AC10"), 1)

    'Negative or PE reviews
    Case "MA_Neg", "MA_PE_"
    'Today Date
        datasourcews.Range("J1") = Date
    'Client Name
        datasourcews.Range("A5") = StrConv(thisws.Range("B7"), vbProperCase)
    'reformat street address
        If thisws.Range("L7") <> "" Then
            datasourcews.Range("A6") = StrConv(thisws.Range("L6"), vbProperCase) & ", " & StrConv(thisws.Range("L7"), vbProperCase)
        Else
            datasourcews.Range("A6") = StrConv(thisws.Range("L6"), vbProperCase)
        End If
    'Reformat city and state
        temparry = Split(thisws.Range("L8"), ",")
        TempStr = StrConv(Trim(temparry(LBound(temparry))), vbProperCase)
        datasourcews.Range("A7") = TempStr & ", " & Trim(temparry(UBound(temparry)))
    'county/case number
        datasourcews.Range("D5") = Right(thisws.Range("G15"), 2) & "/" & Left(thisws.Range("S15"), 9)
    'county number - hidden on sheet
        datasourcews.Range("D8") = Right(thisws.Range("G15"), 2)
    'district number - hidden on sheet
        datasourcews.Range("E8") = thisws.Range("J15") & thisws.Range("K15")
    'review number
        datasourcews.Range("F5") = thisws.Range("L15")
        review_number = datasourcews.Range("F5")
    'review month
        datasourcews.Range("H5") = Left(thisws.Range("T11"), 2) & "/" & Right(thisws.Range("T11"), 4)
        sample_month = thisws.Range("AM10")
    'County Name
        datasourcews.Range("J6") = thisws.Range("F2")
    
    'use ds worksheet to find out qc manager, supervisor and program manager
        DSws.Range("A7") = Val(thisws.Range("AB11") & thisws.Range("AC11")) 'examiner id
        reviewerid = DSws.Range("A7")
        DSws.Range("G2") = DSws.Range("C7") & "/" & DSws.Range("AM2") 'supervisor/examiner name
        DSws.Range("F7") = review_type
        DSws.Range("E7") = review_type

'put in county and district to get email address
    DSws.Range("A49") = datasourcews.Range("D8")
    If datasourcews.Range("E8") = "" Then
        DSws.Range("B49") = "BB"
    Else
        DSws.Range("B49") = WorksheetFunction.Text(datasourcews.Range("E8"), "00")
    End If
'cao email - hidden on sheet
    datasourcews.Range("F8") = DSws.Range("C49")
    
    'fill in the above information in finding memo
    datasourcews.Range("A10") = DSws.Range("H2")  ' Program Manager
    datasourcews.Range("D10") = DSws.Range("AP2") ' Supervisor
    datasourcews.Range("I10") = DSws.Range("AM2") ' Examiner
    
    'close without saving DS file and then delete it
        DSwb.Close False
        Kill tempsourcenameDS
    
    ' Get county number and name in correct format
        countynum = Right(thisws.Range("G15"), 2)
        districtnum = thisws.Range("J15") & thisws.Range("K15")
        countylookup = countyformat(countynum, thisws.Range("F2"), districtnum)
        areanum = Left(thisws.Range("G15"), 1)
    
End Select
    
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////
    
'what type of Finding memo
Select Case memo_type
    'MA Positive Finding
    Case "MA_Pos_Find"
        'Loop through first four rows of Elements, Nature and Cause codes
        TempStr = ""
        For irow = 96 To 102 Step 2
            'if this row contains an element, then add codes to the tempstr
            If thisws.Range("O" & irow) <> "" Then
                If TempStr = "" Then
                    TempStr = thisws.Range("O" & irow) & " - " & thisws.Range("T" & irow) & " - " & thisws.Range("X" & irow)
                Else
                    TempStr = TempStr & vbNewLine & thisws.Range("O" & irow) & " - " & thisws.Range("T" & irow) & " - " & thisws.Range("X" & irow)
                End If
            Else
                Exit For
            End If
        Next irow
        
        'insert codes into findings memo
        If TempStr <> "" Then
            datasourcews.Range("H12") = TempStr
        End If

    'MA Negative Finding
    Case "MA_Neg_Find"
        'Loop through first four rows of Elements and Nature Codes
        TempStr = ""
        
        'check if section H has a 2 and then add codes to the tempstr
        If thisws.Range("C25") = 2 Then
            TempStr = "H - 2"
        End If
        'check if section I. (b) has a 2 and then add codes to the tempstr
        If thisws.Range("C40") = 2 Then
            If TempStr = "" Then
                TempStr = "I. (b) - 2"
            Else
                TempStr = TempStr & vbNewLine & "I. (b) - 2"
            End If
        End If
        
        'insert codes into findings memo
        If TempStr <> "" Then
            datasourcews.Range("H12") = TempStr
        End If
        
End Select

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Add CC list to memo
'Open CC list file
    Workbooks.Open fileName:=PathStr & "Finding Memo\Courtesy Copy for Memos.xlsx", UpdateLinks:=False, ReadOnly:=True
    Set ccwb = ActiveWorkbook
    Set ccws = ccwb.Worksheets(areanum)

'Form cc string
    ccstring = vbNewLine & vbNewLine & "cc: "
    ccstring = ccstring & ccws.Range("B4") & vbNewLine 'Area Manager
    ccstring = ccstring & ccws.Range("B5") & vbNewLine 'Area Staff Asst
    ccrow = 6
    'check area number
    If areanum = "1" Then
        ccstring = ccstring & ccws.Range("B" & ccrow) & vbNewLine 'QA for Area 1
        ccrow = ccrow + 1
    End If
    
    ccstring = ccstring & datasourcews.Range("A10") & vbNewLine 'Office manager
    ccrow = ccrow + 1
    
    'Column B is MA Findings column in Courtesy Copy for Memos file
    cc_col = "B"
    
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
        
'Close cc workbook
    ccwb.Close False
' Add cc string to textbox
    datasourcews.Shapes("TextBox 1").TextFrame.Characters.Text = datasourcews.Shapes("TextBox 1").TextFrame.Characters.Text & ccstring

'Format top line as bold underline - different if finding or deficiency memo
    If InStr(memo_type, "Find") > 0 Then
        datasourcews.Shapes.Range(Array("TextBox 1")).Select
        Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 24).Font.Bold = _
                msoTrue
        Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 24).Font. _
                UnderlineStyle = msoUnderlineSingleLine
    Else
        datasourcews.Shapes.Range(Array("TextBox 1")).Select
        Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 26).Font.Bold = _
                msoTrue
        Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 26).Font. _
                UnderlineStyle = msoUnderlineSingleLine
    End If

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Going out to website with list of Executive Directors
        'FullTextFileName = DSws.Range("U6")
            
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
                    datasourcews.Range("J5") = Range("D" & irow)
                    Exit For
                End If
            Next irow
            ActiveWorkbook.Close False
            
    findmemo = final_title & review_number & " for Sample Month " & sample_month & ".xlsm"
        
    'Check to see if findings memo has already been created.
    If Dir(sPath & "\" & findmemo) <> "" Then
        Response = MsgBox("This memo has already been created. Do you want to overwrite this memo? No memo will be created if you click NO", Buttons:=vbYesNo)
    ' If statement to check if the NO button was selected.
      If Response = vbNo Then
        End
      Else
      'if finding memo exists, then delete it and create new memo
        Kill sPath & "\" & findmemo
      End If
    End If
    
'save finding memo and then rename it

datasourcewb.Close True
Name tempsourcename As sPath & "\" & findmemo

Application.StatusBar = False
Application.ScreenUpdating = True
Application.Visible = True

MsgBox "A " & final_title & " has been saved as " & sPath & "\" & findmemo & "."

End Sub

Function find_texts_fs(fsworkbook_ws As Worksheet)

    'initialize text string to heading at top of text box
    find_texts = "Summary of the Finding:" & vbNewLine & vbNewLine

    'look if buttons labeled 2 or 3 in first element were selected
    If fsworkbook_ws.Shapes("OB 59").OLEFormat.Object.Value = xlOn Or fsworkbook_ws.Shapes("OB 60").OLEFormat.Object.Value = xlOn Then
        find_texts = find_texts & "110 BASIC PROGRAM REQUIREMENTS" & vbNewLine & vbNewLine & vbNewLine
    End If

    'look if buttons labeled 2 or 3 in second element were selected and etc
    If fsworkbook_ws.Shapes("OB 65").OLEFormat.Object.Value = xlOn Or fsworkbook_ws.Shapes("OB 66").OLEFormat.Object.Value = xlOn Then
        find_texts = find_texts & "111 STUDENT STATUS" & vbNewLine & vbNewLine & vbNewLine
    End If

    If fsworkbook_ws.Shapes("OB 69").OLEFormat.Object.Value = xlOn Or fsworkbook_ws.Shapes("OB 70").OLEFormat.Object.Value = xlOn Then
        find_texts = find_texts & "130 CITIZENSHIP AND ALIENAGE" & vbNewLine & vbNewLine & vbNewLine
    End If

    If fsworkbook_ws.Shapes("OB 73").OLEFormat.Object.Value = xlOn Or fsworkbook_ws.Shapes("OB 74").OLEFormat.Object.Value = xlOn Then
        find_texts = find_texts & "140 RESIDENCY" & vbNewLine & vbNewLine & vbNewLine
    End If

    If fsworkbook_ws.Shapes("OB 90").OLEFormat.Object.Value = xlOn Or fsworkbook_ws.Shapes("OB 91").OLEFormat.Object.Value = xlOn Then
        find_texts = find_texts & "150 HOUSEHOLD COMPOSITION" & vbNewLine & vbNewLine & vbNewLine
    End If

    If fsworkbook_ws.Shapes("OB 113").OLEFormat.Object.Value = xlOn Or fsworkbook_ws.Shapes("OB 114").OLEFormat.Object.Value = xlOn Then
        find_texts = find_texts & "151 RECIPIENT DISQUALIFICATION" & vbNewLine & vbNewLine & vbNewLine
    End If
    
    If fsworkbook_ws.Shapes("OB 134").OLEFormat.Object.Value = xlOn Or fsworkbook_ws.Shapes("OB 135").OLEFormat.Object.Value = xlOn Then
        find_texts = find_texts & "160 EMPLOYMENT AND TRAINING PROGRAMS" & vbNewLine & vbNewLine & vbNewLine
    End If
    
    If fsworkbook_ws.Shapes("OB 213").OLEFormat.Object.Value = xlOn Or fsworkbook_ws.Shapes("OB 156").OLEFormat.Object.Value = xlOn Then
        find_texts = find_texts & "161 TIME LIMIT PARTICIPATION" & vbNewLine & vbNewLine & vbNewLine
    End If
    
    If fsworkbook_ws.Shapes("OB 175").OLEFormat.Object.Value = xlOn Or fsworkbook_ws.Shapes("OB 176").OLEFormat.Object.Value = xlOn Then
        find_texts = find_texts & "162 WORK REGISTRATION" & vbNewLine & vbNewLine & vbNewLine
    End If
    
    If fsworkbook_ws.Shapes("OB 267").OLEFormat.Object.Value = xlOn Or fsworkbook_ws.Shapes("OB 268").OLEFormat.Object.Value = xlOn Then
        find_texts = find_texts & "163 VOLUNTARY QUIT/REDUCING WORK EFFORT" & vbNewLine & vbNewLine & vbNewLine
    End If
    
    If fsworkbook_ws.Shapes("OB 272").OLEFormat.Object.Value = xlOn Or fsworkbook_ws.Shapes("OB 274").OLEFormat.Object.Value = xlOn Then
        find_texts = find_texts & "164 WORKFARE AND COMPARABLE WORKFARE" & vbNewLine & vbNewLine & vbNewLine
    End If
    
    If fsworkbook_ws.Shapes("OB 221").OLEFormat.Object.Value = xlOn Or fsworkbook_ws.Shapes("OB 222").OLEFormat.Object.Value = xlOn Then
        find_texts = find_texts & "165 EMPLOYMENT STATUS/JOB AVAILABILITY" & vbNewLine & vbNewLine & vbNewLine
    End If
    
    If fsworkbook_ws.Shapes("OB 263").OLEFormat.Object.Value = xlOn Or fsworkbook_ws.Shapes("OB 262").OLEFormat.Object.Value = xlOn Then
        find_texts = find_texts & "166 ACCEPTANCE OF EMPLOYMENT" & vbNewLine & vbNewLine & vbNewLine
    End If
    
    If fsworkbook_ws.Shapes("OB 280").OLEFormat.Object.Value = xlOn Or fsworkbook_ws.Shapes("OB 279").OLEFormat.Object.Value = xlOn Then
        find_texts = find_texts & "170 SOCIAL SECURITY NUMBER" & vbNewLine & vbNewLine & vbNewLine
    End If
    
    If fsworkbook_ws.Shapes("OB 584").OLEFormat.Object.Value = xlOn Or fsworkbook_ws.Shapes("OB 585").OLEFormat.Object.Value = xlOn Then
        find_texts = find_texts & "211 BANK ACCOUNTS OR CASH ON HAND" & vbNewLine & vbNewLine & vbNewLine
    End If
    
    If fsworkbook_ws.Shapes("OB 543").OLEFormat.Object.Value = xlOn Or fsworkbook_ws.Shapes("OB 544").OLEFormat.Object.Value = xlOn Then
        find_texts = find_texts & "212 NON RECURRING LUMP SUM PAYMENTS" & vbNewLine & vbNewLine & vbNewLine
    End If
    
    If fsworkbook_ws.Shapes("OB 540").OLEFormat.Object.Value = xlOn Or fsworkbook_ws.Shapes("OB 468").OLEFormat.Object.Value = xlOn Then
        find_texts = find_texts & "213 OTHER LIQUID ASSETS" & vbNewLine & vbNewLine & vbNewLine
    End If
    
    If fsworkbook_ws.Shapes("OB 469").OLEFormat.Object.Value = xlOn Or fsworkbook_ws.Shapes("OB 471").OLEFormat.Object.Value = xlOn Then
        find_texts = find_texts & "221 REAL PROPERTY" & vbNewLine & vbNewLine & vbNewLine
    End If
    
    If fsworkbook_ws.Shapes("OB 473").OLEFormat.Object.Value = xlOn Or fsworkbook_ws.Shapes("OB 472").OLEFormat.Object.Value = xlOn Then
        find_texts = find_texts & "222 VEHICLE" & vbNewLine & vbNewLine & vbNewLine
    End If
    
    If fsworkbook_ws.Shapes("OB 549").OLEFormat.Object.Value = xlOn Or fsworkbook_ws.Shapes("OB 550").OLEFormat.Object.Value = xlOn Then
        find_texts = find_texts & "224 OTHER NON LIQUID RESOURCES" & vbNewLine & vbNewLine & vbNewLine
    End If
    
    If fsworkbook_ws.Shapes("OB 554").OLEFormat.Object.Value = xlOn Or fsworkbook_ws.Shapes("OB 553").OLEFormat.Object.Value = xlOn Then
        find_texts = find_texts & "225 COMBINED RESOURCES" & vbNewLine & vbNewLine & vbNewLine
    End If
    
    If fsworkbook_ws.Shapes("OB 557").OLEFormat.Object.Value = xlOn Or fsworkbook_ws.Shapes("OB 558").OLEFormat.Object.Value = xlOn Then
        find_texts = find_texts & "311 WAGES AND SALARIES" & vbNewLine & vbNewLine & vbNewLine
    End If
    
    If fsworkbook_ws.Shapes("OB 560").OLEFormat.Object.Value = xlOn Or fsworkbook_ws.Shapes("OB 561").OLEFormat.Object.Value = xlOn Then
        find_texts = find_texts & "312 SELF EMPLOYMENT" & vbNewLine & vbNewLine & vbNewLine
    End If
    
    If fsworkbook_ws.Shapes("OB 527").OLEFormat.Object.Value = xlOn Or fsworkbook_ws.Shapes("OB 518").OLEFormat.Object.Value = xlOn Then
        find_texts = find_texts & "313 EARNED INCOME CREDIT" & vbNewLine & vbNewLine & vbNewLine
    End If
    
    If fsworkbook_ws.Shapes("OB 525").OLEFormat.Object.Value = xlOn Or fsworkbook_ws.Shapes("OB 526").OLEFormat.Object.Value = xlOn Then
        find_texts = find_texts & "314 OTHER EARNED INCOME" & vbNewLine & vbNewLine & vbNewLine
    End If
    
    If fsworkbook_ws.Shapes("OB 637").OLEFormat.Object.Value = xlOn Or fsworkbook_ws.Shapes("OB 636").OLEFormat.Object.Value = xlOn Then
        find_texts = find_texts & "321 EARNED INCOME DEDUCTIONS" & vbNewLine & vbNewLine & vbNewLine
    End If
    
    If fsworkbook_ws.Shapes("OB 640").OLEFormat.Object.Value = xlOn Or fsworkbook_ws.Shapes("OB 639").OLEFormat.Object.Value = xlOn Then
        find_texts = find_texts & "323 CHILD OR DEPENDENT CARE" & vbNewLine & vbNewLine & vbNewLine
    End If
    
    If fsworkbook_ws.Shapes("OB 632").OLEFormat.Object.Value = xlOn Or fsworkbook_ws.Shapes("OB 633").OLEFormat.Object.Value = xlOn Then
        find_texts = find_texts & "331 RSDI BENEFITS" & vbNewLine & vbNewLine & vbNewLine
    End If
    
    If fsworkbook_ws.Shapes("OB 668").OLEFormat.Object.Value = xlOn Or fsworkbook_ws.Shapes("OB 667").OLEFormat.Object.Value = xlOn Then
        find_texts = find_texts & "332 VETERANS BENEFITS" & vbNewLine & vbNewLine & vbNewLine
    End If
    
    If fsworkbook_ws.Shapes("OB 682").OLEFormat.Object.Value = xlOn Or fsworkbook_ws.Shapes("OB 681").OLEFormat.Object.Value = xlOn Then
        find_texts = find_texts & "333 SSI AND/OR STATE SSI SUPPLEMENT (SSP)" & vbNewLine & vbNewLine & vbNewLine
    End If
    
    If fsworkbook_ws.Shapes("OB 697").OLEFormat.Object.Value = xlOn Or fsworkbook_ws.Shapes("OB 696").OLEFormat.Object.Value = xlOn Then
        find_texts = find_texts & "334 UNEMPLOYMENT COMPENSATION" & vbNewLine & vbNewLine & vbNewLine
    End If
    
    If fsworkbook_ws.Shapes("OB 710").OLEFormat.Object.Value = xlOn Or fsworkbook_ws.Shapes("OB 709").OLEFormat.Object.Value = xlOn Then
        find_texts = find_texts & "335 WORKERS COMPENSATION" & vbNewLine & vbNewLine & vbNewLine
    End If
    
    If fsworkbook_ws.Shapes("OB 856").OLEFormat.Object.Value = xlOn Or fsworkbook_ws.Shapes("OB 723").OLEFormat.Object.Value = xlOn Then
        find_texts = find_texts & "336 OTHER GOVERNMENT BENEFITS" & vbNewLine & vbNewLine & vbNewLine
    End If
    
    If fsworkbook_ws.Shapes("OB 737").OLEFormat.Object.Value = xlOn Or fsworkbook_ws.Shapes("OB 736").OLEFormat.Object.Value = xlOn Then
        find_texts = find_texts & "342 CONTRIBUTIONS, INCOME-IN-KIND" & vbNewLine & vbNewLine & vbNewLine
    End If
    
    If fsworkbook_ws.Shapes("OB 755").OLEFormat.Object.Value = xlOn Or fsworkbook_ws.Shapes("OB 754").OLEFormat.Object.Value = xlOn Then
        find_texts = find_texts & "343 DEEMED INCOME" & vbNewLine & vbNewLine & vbNewLine
    End If
    
    If fsworkbook_ws.Shapes("OB 774").OLEFormat.Object.Value = xlOn Or fsworkbook_ws.Shapes("OB 773").OLEFormat.Object.Value = xlOn Then
        find_texts = find_texts & "344 GA" & vbNewLine & vbNewLine & vbNewLine
    End If
    
    If fsworkbook_ws.Shapes("OB 790").OLEFormat.Object.Value = xlOn Or fsworkbook_ws.Shapes("OB 789").OLEFormat.Object.Value = xlOn Then
        find_texts = find_texts & "345 EDUCATIONAL GRANTS/SCHOLARSHIPS LOANS" & vbNewLine & vbNewLine & vbNewLine
    End If
    
    If fsworkbook_ws.Shapes("OB 804").OLEFormat.Object.Value = xlOn Or fsworkbook_ws.Shapes("OB 803").OLEFormat.Object.Value = xlOn Then
        find_texts = find_texts & "346 OTHER" & vbNewLine & vbNewLine & vbNewLine
    End If
    
    If fsworkbook_ws.Shapes("OB 818").OLEFormat.Object.Value = xlOn Or fsworkbook_ws.Shapes("OB 817").OLEFormat.Object.Value = xlOn Then
        find_texts = find_texts & "347 TANF" & vbNewLine & vbNewLine & vbNewLine
    End If
    
    If fsworkbook_ws.Shapes("OB 839").OLEFormat.Object.Value = xlOn Or fsworkbook_ws.Shapes("OB 838").OLEFormat.Object.Value = xlOn Then
        find_texts = find_texts & "350 CHILD SUPPORT PAYMENTS RECEIVED FROM ABSENT PARENT" & vbNewLine & vbNewLine & vbNewLine
    End If
    
    If fsworkbook_ws.Shapes("OB 848").OLEFormat.Object.Value = xlOn Or fsworkbook_ws.Shapes("OB 849").OLEFormat.Object.Value = xlOn Then
        find_texts = find_texts & "361 OTHER DISREGARD/DEDUCTIONS" & vbNewLine & vbNewLine & vbNewLine
    End If
    
    If fsworkbook_ws.Shapes("OB 876").OLEFormat.Object.Value = xlOn Or fsworkbook_ws.Shapes("OB 877").OLEFormat.Object.Value = xlOn Then
        find_texts = find_texts & "363 SHELTER DEDUCTION" & vbNewLine & vbNewLine & vbNewLine
    End If
    
    If fsworkbook_ws.Shapes("OB 1024").OLEFormat.Object.Value = xlOn Or fsworkbook_ws.Shapes("OB 1023").OLEFormat.Object.Value = xlOn Then
        find_texts = find_texts & "364 STANDARD UTILITY ALLOWANCE" & vbNewLine & vbNewLine & vbNewLine
    End If
    
    If fsworkbook_ws.Shapes("OB 901").OLEFormat.Object.Value = xlOn Or fsworkbook_ws.Shapes("OB 900").OLEFormat.Object.Value = xlOn Then
        find_texts = find_texts & "365 MEDICAL DEDUCTIONS" & vbNewLine & vbNewLine & vbNewLine
    End If
    
    If fsworkbook_ws.Shapes("OB 917").OLEFormat.Object.Value = xlOn Or fsworkbook_ws.Shapes("OB 916").OLEFormat.Object.Value = xlOn Then
        find_texts = find_texts & "366 CHILD SUPPORT OBLIGATION" & vbNewLine & vbNewLine & vbNewLine
    End If
    
    If fsworkbook_ws.Shapes("OB 933").OLEFormat.Object.Value = xlOn Or fsworkbook_ws.Shapes("OB 982").OLEFormat.Object.Value = xlOn Then
        find_texts = find_texts & "371 COMBINED GROSS INCOME" & vbNewLine & vbNewLine & vbNewLine
    End If
    
    If fsworkbook_ws.Shapes("OB 949").OLEFormat.Object.Value = xlOn Or fsworkbook_ws.Shapes("OB 948").OLEFormat.Object.Value = xlOn Then
        find_texts = find_texts & "372 COMBINED NET INCOME" & vbNewLine & vbNewLine & vbNewLine
    End If
    
    If fsworkbook_ws.Shapes("OB 959").OLEFormat.Object.Value = xlOn Or fsworkbook_ws.Shapes("OB 975").OLEFormat.Object.Value = xlOn Then
        find_texts = find_texts & "520 ARITHMETIC" & vbNewLine & vbNewLine & vbNewLine
    End If
    
    If fsworkbook_ws.Shapes("OB 969").OLEFormat.Object.Value = xlOn Or fsworkbook_ws.Shapes("OB 968").OLEFormat.Object.Value = xlOn Then
        find_texts = find_texts & "560 SEMI-ANNUALLY REPORTING" & vbNewLine & vbNewLine & vbNewLine
    End If
    
    find_texts_fs = find_texts

End Function

Function find_texts_tanf(fsworkbook_ws As Worksheet)

    'initialize text string to heading at top of text box
    find_texts = vbNewLine & vbNewLine

    'look if buttons labeled 2 or 3 in first element were selected
    If fsworkbook_ws.Shapes("OB 89").OLEFormat.Object.Value = xlOn Or fsworkbook_ws.Shapes("OB 90").OLEFormat.Object.Value = xlOn Then
        find_texts = find_texts & "151 DEFINITIVE CONDITIONS PART 1" & vbNewLine & vbNewLine & vbNewLine
    End If

    'look if buttons labeled 2 or 3 in second element were selected and etc
    If fsworkbook_ws.Shapes("OB 193").OLEFormat.Object.Value = xlOn Or fsworkbook_ws.Shapes("OB 194").OLEFormat.Object.Value = xlOn Then
        find_texts = find_texts & "151 DEFINITIVE CONDITIONS PART 2" & vbNewLine & vbNewLine & vbNewLine
    End If

    If fsworkbook_ws.Shapes("OB 299").OLEFormat.Object.Value = xlOn Or fsworkbook_ws.Shapes("OB 300").OLEFormat.Object.Value = xlOn Then
        find_texts = find_texts & "151 DEFINITIVE CONDITIONS PART 3" & vbNewLine & vbNewLine & vbNewLine
    End If

    'If fsworkbook_ws.Shapes("OB 254").OLEFormat.Object.value = xlOn Or fsworkbook_ws.Shapes("OB 255").OLEFormat.Object.value = xlOn Then
     '   find_texts = find_texts & "160 RESET" & vbNewLine & vbNewLine & vbNewLine
    'End If

    If fsworkbook_ws.Shapes("OB 257").OLEFormat.Object.Value = xlOn Or fsworkbook_ws.Shapes("OB 258").OLEFormat.Object.Value = xlOn Then
        find_texts = find_texts & "164 AMR" & vbNewLine & vbNewLine & vbNewLine
    End If

    If fsworkbook_ws.Shapes("OB 261").OLEFormat.Object.Value = xlOn Or fsworkbook_ws.Shapes("OB 262").OLEFormat.Object.Value = xlOn Then
        find_texts = find_texts & "164 SSN" & vbNewLine & vbNewLine & vbNewLine
    End If
    
    If fsworkbook_ws.Shapes("OB 310").OLEFormat.Object.Value = xlOn Or fsworkbook_ws.Shapes("OB 311").OLEFormat.Object.Value = xlOn Then
        find_texts = find_texts & "189 CRIMINAL HISTORY" & vbNewLine & vbNewLine & vbNewLine
    End If
    
    If fsworkbook_ws.Shapes("OB 325").OLEFormat.Object.Value = xlOn Or fsworkbook_ws.Shapes("OB 326").OLEFormat.Object.Value = xlOn Then
        find_texts = find_texts & "311 WAGES SALARIES DEEMED INCOME" & vbNewLine & vbNewLine & vbNewLine
    End If
    
    If fsworkbook_ws.Shapes("OB 376").OLEFormat.Object.Value = xlOn Or fsworkbook_ws.Shapes("OB 377").OLEFormat.Object.Value = xlOn Then
        find_texts = find_texts & "334 UNEMPLOYMENT COMPENSATION" & vbNewLine & vbNewLine & vbNewLine
    End If
    
    If fsworkbook_ws.Shapes("OB 389").OLEFormat.Object.Value = xlOn Or fsworkbook_ws.Shapes("OB 390").OLEFormat.Object.Value = xlOn Then
        find_texts = find_texts & "520 ARITHMETIC COMPUTATION" & vbNewLine & vbNewLine & vbNewLine
    End If
    
    If fsworkbook_ws.Shapes("OB 494").OLEFormat.Object.Value = xlOn Or fsworkbook_ws.Shapes("OB 495").OLEFormat.Object.Value = xlOn Then
        find_texts = find_texts & "420 SPECIAL ALLOWANCES PART 1" & vbNewLine & vbNewLine & vbNewLine
    End If
    
  '  If fsworkbook_ws.Shapes("OB 530").OLEFormat.Object.value = xlOn Or fsworkbook_ws.Shapes("OB 539").OLEFormat.Object.value = xlOn Then
  '      find_texts = find_texts & "420 SPECIAL ALLOWANCES PART 2" & vbNewLine & vbNewLine & vbNewLine
  '  End If

    find_texts_tanf = find_texts

End Function

Sub MailMergeandSave()

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

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////
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
Range("F7") = review_type
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
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
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
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'TANF
ElseIf review_type = 1 Then

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
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'PE review
ElseIf Left(thisws.Name, 2) = "24" Then
review_type = 9
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
sample_month = Range("E2")
'reviewer number
Range("A7") = Val(thisws.Range("AB11") & thisws.Range("AC11"))
'supervisor/examiner
Range("G2") = Range("C7") & "/" & Range("BA2")
'Range("R2") = thisws.Range("J61") & " - " & thisws.Range("O61")
'finding
Range("E17") = Val(thisws.Range("AL10"))
Range("F7") = review_type
Range("E7") = review_type


'MA
ElseIf review_type = 2 Then

countynum = Right(thisws.Range("AC10"), 2)
districtnum = thisws.Range("AH10")
countylookup = countyformat(countynum, thisws.Range("O4"), districtnum)

'Client Information
Range("B2") = StrConv(thisws.Range("B2"), vbProperCase)
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

'Case Number
Range("C2") = Right(thisws.Range("AC10"), 2) & "/" & Left(thisws.Range("H10"), 9)

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
Range("R2") = thisws.Range("J61") & " - " & thisws.Range("O61")
'finding
Range("E17") = Val(thisws.Range("AL10"))
Range("F7") = review_type
Range("E7") = review_type
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'MA Negative
ElseIf review_type = 8 Then
countynum = Right(thisws.Range("G15"), 2)
districtnum = thisws.Range("J15") & thisws.Range("K15")
countylookup = countyformat(countynum, thisws.Range("F2"), districtnum)
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
sample_month = Range("E2")
'reviewer number
Range("A7") = Val(thisws.Range("AB11") & thisws.Range("AC11"))
'supervisor/examiner
Range("G2") = Range("C7") & "/" & Range("BA2")
'Range("R2") = thisws.Range("J61") & " - " & thisws.Range("O61")
'finding
Range("E17") = Val(thisws.Range("AL10"))
Range("F7") = review_type
Range("E7") = review_type


End If
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
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
            
    If review_type = 2 Then
        sample_month = Left(sample_month, 2) & Right(sample_month, 4)
        findmemo = "QC FINDING Review Number " & review_number & " for Sample Month " & sample_month & ".doc"
    Else
        findmemo = "QC FINDING Review Number " & review_number & " for Sample Month " & sample_month & ".doc"
    End If
    
    'Copy template finding memo to active directory
    SrceFile = PathStr & "Finding Memo\QC FINDING MASTER Revised 3-08.doc"
    tempfindmemoname = sPath & "\FM temp.doc"
    DestFile = tempfindmemoname
    FileCopy SrceFile, DestFile
    
    'Check to see if findings memo has already been created.
    If Dir(sPath & "\" & findmemo) <> "" Then
        Response = MsgBox("A findings memo has already been created. Do you want to overwrite this memo? No memo will be created if you click NO", Buttons:=vbYesNo)
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

MsgBox "A Finding Memo has been saved as " & sPath & "\" & findmemo & _
". Please close the Internet Explorer with the CAO Directory. - Thank you!"

End Sub
Sub PotentialErrorCall()

'Create Potential Error Call Memo
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
Range("CB2") = thisws.Range("L16") 'most recent action
Range("CC2") = thisws.Range("AH20") 'income disregard
Range("CD2") = thisws.Range("I20") 'sample month payment

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
Range("CB2") = thisws.Range("L16") 'most recent action
Range("CC2") = thisws.Range("AJ20") 'income disregard
Range("CD2") = thisws.Range("G20") 'sample month payment

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'MA Positive
ElseIf review_type = 2 Then

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

'Client Information
Range("B2") = StrConv(thisws.Range("B2"), vbProperCase)
'Case Number
Range("C2") = Right(thisws.Range("AC10"), 2) & "/" & Left(thisws.Range("H10"), 9)

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
Range("R2") = thisws.Range("J61") & " - " & thisws.Range("O61")
Range("F7") = review_type
Range("E7") = review_type
Range("CE2") = thisws.Range("Q10")
Range("CF2") = "POSITIVE"

'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'MA Negative
ElseIf review_type = 8 Then

' Get county number and name in correct format
countynum = Right(thisws.Range("G15"), 2)
districtnum = thisws.Range("J15") & thisws.Range("K15")
countylookup = countyformat(countynum, thisws.Range("F2"), districtnum)
Range("AK2") = thisws.Range("M5") & " Assistance Office"
'Look for counties with districts
'If countynum = "02" Or countynum = "23" Or countynum = "40" Or _
'        countynum = "51" Or countynum = "63" Or countynum = "65" Then
'    Range("AK2") = thisws.Range("F2") & " Office"
'Else
'    Range("AK2") = thisws.Range("F2") & " CAO"
'End If

'Client Information
Range("B2") = StrConv(thisws.Range("B7"), vbProperCase)

'Case Number
Range("C2") = Right(thisws.Range("G15"), 2) & "/" & thisws.Range("S15")

'review number
Range("D2") = thisws.Range("L15")
review_number = Range("D2")
'review month
Range("E2") = thisws.Range("T11")
sample_month = Range("E2")
'reviewer number
Range("A7") = Val(thisws.Range("AB11") & thisws.Range("AC11"))
'supervisor/examiner
Range("G2") = Range("C7") & "/" & Range("BA2")
'Range("R2") = thisws.Range("J61") & " - " & thisws.Range("O61")
Range("F7") = review_type
Range("E7") = review_type
Range("CE2") = thisws.Range("AB15")
Range("CF2") = "NEGATIVE"

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
            
    If review_type = 1 Then
        findmemo = "TANF Case Summary Review Number " & review_number & " for Sample Month " & sample_month & ".docx"
        SrceFile = PathStr & "Finding Memo\TANF PEC Case Summary.docx"
    ElseIf review_type = 9 Then
        findmemo = "GA Case Summary Review Number " & review_number & " for Sample Month " & sample_month & ".docx"
        SrceFile = PathStr & "Finding Memo\GA Case Summary.docx"
    ElseIf review_type = 2 Then
        sample_month = Left(sample_month, 2) & Right(sample_month, 4)
        findmemo = "MA Potential Positive Error Summary Review Number " & review_number & " for Sample Month " & sample_month & ".docx"
        SrceFile = PathStr & "Finding Memo\MA Potential Error Summary.docx"
    ElseIf review_type = 8 Then
        findmemo = "MA Potential Negative Error Summary Review Number " & review_number & " for Sample Month " & sample_month & ".docx"
        SrceFile = PathStr & "Finding Memo\MA Potential Error Summary.docx"
    End If
    
    'Copy template finding memo to active directory
    tempfindmemoname = sPath & "\FM temp.docx"
    DestFile = tempfindmemoname
    FileCopy SrceFile, DestFile
    
    'Check to see if findings memo has already been created.
    If Dir(sPath & "\" & findmemo) <> "" Then
        Response = MsgBox("A potential error memo has already been created. Do you want to overwrite this memo? No memo will be created if you click NO", Buttons:=vbYesNo)
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

MsgBox "A Potential Error Memo has been saved as " & sPath & "\" & findmemo

End Sub
Sub PE_Deficiency_Letter()

'Create PE Letter
Dim appWD As Object
Dim review_number As Long
Dim sample_month As String, PathStr As String, findmemo As String
Dim countynum As String, districtnum As String

'Find path to Examiner's files on this PC
Set WshNetwork = CreateObject("WScript.Network")
Set oDrives = WshNetwork.EnumNetworkDrives

Dim DLetter As String, DUNC As String
Dim thisws As Worksheet, thiswb As Workbook
Dim providerws As Worksheet, providerwb As Workbook
Dim datasourcewb As Workbook, datasourcews As Worksheet
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

Set datasourcewb = ActiveWorkbook
Set datasourcews = ActiveSheet

'deleting information in Finding Memo Data Source spreadsheet
datasourcews.Range("B2:E2") = ""
datasourcews.Range("G2") = ""
datasourcews.Range("O2:S2") = ""
datasourcews.Range("W2:Y2") = ""

'Case Information
datasourcews.Range("B2") = StrConv(thisws.Range("B7"), vbProperCase) 'Client Name
datasourcews.Range("D2") = thisws.Range("L15") 'Review Number
review_number = thisws.Range("L15") 'Review Number

'Get provider information
'Open spreadsheet where provider address
provider_file = PathStr & "Finding Memo\ACA PE Providers Contact 5-2-14.xlsx"
Workbooks.Open fileName:=provider_file, UpdateLinks:=False, ReadOnly:=True
Set providerwb = ActiveWorkbook
Set providerws = ActiveSheet

'find max row in provider worksheet
maxrow = providerws.Cells.Find(What:="*", _
            SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row

'look for provider number
provider_number = thisws.Range("AB18")
For i = 2 To maxrow
    If providerws.Range("A" & i) = provider_number Then
        datasourcews.Range("CT2") = providerws.Range("D" & i) 'Provider Contact Name
        datasourcews.Range("CU2") = StrConv(providerws.Range("C" & i), vbProperCase) 'Provider Name
        datasourcews.Range("CV2") = StrConv(providerws.Range("G" & i), vbProperCase) 'Provider Street Address
        datasourcews.Range("CW2") = StrConv(providerws.Range("H" & i), vbProperCase) & ", " & _
                                    providerws.Range("I" & i) & " " & _
                                    Right(providerws.Range("J" & i), 5) 'Provider City State Zip
        Exit For
    End If
Next i

If datasourcews.Range("CT2") = "" Then
    MsgBox "No provider has been found with the provider number " & provider_number
End If

datasourcewb.Close True
providerwb.Close False

sample_month = thisws.Range("T11")
findmemo = "PE Deficiency Letter for Review Number " & review_number & " Sample Month " & sample_month & ".docx"
SrceFile = PathStr & "Finding Memo\PE Finding OMAP.docx"
    
    'Copy template finding memo to active directory
    tempfindmemoname = sPath & "\FM temp.docx"
    DestFile = tempfindmemoname
    FileCopy SrceFile, DestFile
    
    'Check to see if findings memo has already been created.
    If Dir(sPath & "\" & findmemo) <> "" Then
        Response = MsgBox("A PE Deficiency Letter has already been created. Do you want to overwrite this memo? No memo will be created if you click NO", Buttons:=vbYesNo)
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

MsgBox "A PE Deficiency Letter has been saved as " & sPath & "\" & findmemo

End Sub
Function countyformat(countynum As String, countyname As String, districtnum As String) As String
  
  If countynum = "02" Or countynum = "23" Or countynum = "40" Or _
        countynum = "51" Or countynum = "63" Or countynum = "65" Then
    countyformat = " "
    Select Case districtnum
        Case "02"
            countyformat = "02/2"
        Case "03"
            countyformat = "02/7"
        Case "05"
            countyformat = "02/3"
        Case "06"
            countyformat = "02/6"
        Case "07"
            countyformat = "02/9"
        Case "08"
            countyformat = "02/4"
        Case "09"
            countyformat = "02/5"
        Case "10"
            countyformat = "02/A"
        Case "15"
            countyformat = "65/1"
        Case "16"
            countyformat = "65/2"
        Case "17"
            countyformat = "65/4"
        Case "20"
            countyformat = "63/2"
       ' Case "21"
       '     countyformat = "63/1 - Washington "
        Case "21"
            countyformat = "63/1"
        Case "30"
            countyformat = "40/2"
        Case "33"
            countyformat = "40/1"
        Case "41"
            countyformat = "51/F"
        Case "42"
            countyformat = "51/G"
        Case "44"
            countyformat = "51/6"
        Case "45"
            countyformat = "51/D"
        Case "46"
            countyformat = "51/2"
        Case "47"
            countyformat = "51/K"
        Case "48"
            countyformat = "51/9"
        Case "49"
            countyformat = "51/3"
        Case "51"
            countyformat = "51/H"
        Case "55"
            countyformat = "51/M"
        Case "58"
            countyformat = "51/4"
        Case "59"
            countyformat = "51/J"
        Case "62"
            countyformat = "51/7"
        Case "63"
            countyformat = "51/P"
        Case "64"
            countyformat = "51/5"
        Case "65"
            countyformat = "51/3"
        Case "80"
            countyformat = "23/1"
        Case "82"
            countyformat = "23/2"
    End Select
  Else
    temparry = Split(countyname, " ")
    countyformat = countynum & " - " & temparry(LBound(temparry))
 End If
End Function






