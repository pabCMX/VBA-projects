Attribute VB_Name = "repopulate_mod"
Sub repop_modules_forms()

Dim maxrow As Integer, maxrowex As Integer
Dim program As String, exname As String, exnumstr As String
Dim PathStr As String
Dim monthstr As String, reviewtxt As String
Dim thisWBname As Workbook, schedule_wb As Workbook
Dim DUNC As String
Dim fs As clFileSearchModule

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

pathdir = DLetter & "\Schedules by Examiner Number\"

If Len(Dir(pathdir, vbDirectory)) = 0 Then
    MsgBox "Path to Examiner's File: " & pathdir & " does NOT exists!!" & Chr(13) & _
        "Contact Valerie or Nicole"
    End
End If

' Keep Screen from updating
Application.ScreenUpdating = False

' Save Display Status Bar
Application.DisplayStatusBar = True

'Set WshShell = CreateObject("Wscript.Shell")
 
' Save name of the search for files workbook and worksheet
Set thisWBname = ActiveWorkbook
Set thissht = ActiveWorkbook.Sheets("repop")

' Find maximum row for cases
thissht.Range("E1").End(xlDown).Select
maxrow = ActiveCell.Row

' Find maximum row for examiners
thissht.Range("L1").End(xlDown).Select
maxrowex = ActiveCell.Row
    
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
            ". " & pct & "% " & i - 2 & "/" & maxrow - 1 & " done. Please be patient..."
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
    Select Case Left(reviewtxt, 2)
    
        Case "50", "51", "55"
            program = "FS Positive"
        Case "60", "61", "65", "66"
            program = "FS Negative"
        Case "14"
            program = "TANF"
        Case "20", "21"
            program = "MA Positive"
        Case "24"
            program = "MA PE"
        Case "80", "81", "82", "83"
            program = "MA Negative"
        Case Else
            MsgBox "Review Number " & reviewtxt & " is not a known QC Reivew number"
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
    'PathStr = "E:\DQC\Schedules by Examiner Number\Previous Year Schedules\SNAP - FFY 2012\"
    'starttime = Now()
            
    ' Search for file in the path
    With fs
        .NewSearch
        .SearchSubFolders = True
        
         ' Set path name
        .LookIn = PathStr
        .FileType = msoFileTypeExcelWorkbooks
        
        ' Set file name
        'Review Number 50002 Month 200910 Examiner 47.xlsm
        .fileName = "Review Number " & reviewtxt & " Month " & monthstr & _
            " Examiner" & "*.xls*"

            ' If file is found then start the copy process
            If .Execute > 0 Then 'Workbooks in folder

                ' Open workbook
                Workbooks.Open fileName:=.FoundFiles(1), UpdateLinks:=False
                Set schedule_wb = ActiveWorkbook
                    
                'copy macros and forms to workbook
                If program = "TANF" Then
                    CopyModule thisWBname, "CAO_Appointment", schedule_wb
                    CopyModule thisWBname, "CashMemos", schedule_wb
                    CopyModule thisWBname, "Drop", schedule_wb
                    CopyModule thisWBname, "Finding_Memo", schedule_wb
                    CopyModule thisWBname, "Module1", schedule_wb
                    CopyModule thisWBname, "Module3", schedule_wb
                    CopyModule thisWBname, "TANFMod", schedule_wb
                    CopyForm thisWBname, "SelectDate", schedule_wb
                    CopyForm thisWBname, "SelectForms", schedule_wb
                    CopyForm thisWBname, "SelectTime", schedule_wb
                    CopyForm thisWBname, "UserForm1", schedule_wb
                    CopyForm thisWBname, "UserForm2", schedule_wb
                ElseIf program = "FS Positive" Then
                    CopyModule thisWBname, "CAO_Appointment", schedule_wb
                    CopyModule thisWBname, "CashMemos", schedule_wb
                    CopyModule thisWBname, "Drop", schedule_wb
                    CopyModule thisWBname, "Finding_Memo", schedule_wb
                    CopyModule thisWBname, "Module1", schedule_wb
                    CopyModule thisWBname, "Module11", schedule_wb
                    CopyModule thisWBname, "Module3", schedule_wb
                    CopyModule thisWBname, "TANFMod", schedule_wb
                    CopyForm thisWBname, "SelectDate", schedule_wb
                    CopyForm thisWBname, "SelectForms", schedule_wb
                    CopyForm thisWBname, "SelectTime", schedule_wb
                ElseIf program = "FS Negative" Then
                    CopyModule thisWBname, "CAO_Appointment", schedule_wb
                    CopyModule thisWBname, "CashMemos", schedule_wb
                    CopyModule thisWBname, "Finding_Memo", schedule_wb
                    CopyModule thisWBname, "Module1", schedule_wb
                    CopyModule thisWBname, "Module3", schedule_wb
                    CopyModule thisWBname, "TANFMod", schedule_wb
                    CopyForm thisWBname, "SelectDate", schedule_wb
                    CopyForm thisWBname, "SelectForms", schedule_wb
                    CopyForm thisWBname, "SelectTime", schedule_wb
                ElseIf program = "MA Positive" Then
                    CopyModule thisWBname, "CAO_Appointment", schedule_wb
                    CopyModule thisWBname, "CashMemos", schedule_wb
                    CopyModule thisWBname, "Drop", schedule_wb
                    CopyModule thisWBname, "Finding_Memo", schedule_wb
                    CopyModule thisWBname, "MA_Comp_mod", schedule_wb
                    CopyModule thisWBname, "Module1", schedule_wb
                    CopyModule thisWBname, "Module3", schedule_wb
                    CopyModule thisWBname, "TANFMod", schedule_wb
                    CopyForm thisWBname, "MASelectForms", schedule_wb
                    CopyForm thisWBname, "SelectDate", schedule_wb
                    CopyForm thisWBname, "SelectForms", schedule_wb
                    CopyForm thisWBname, "SelectTime", schedule_wb
                    CopyForm thisWBname, "UserFormMAC2", schedule_wb
                    CopyForm thisWBname, "UserFormMAC3", schedule_wb
                ElseIf program = "MA Negative" Then
                    CopyModule thisWBname, "CAO_Appointment", schedule_wb
                    CopyModule thisWBname, "CashMemos", schedule_wb
                    CopyModule thisWBname, "Finding_Memo", schedule_wb
                    CopyModule thisWBname, "Module1", schedule_wb
                    CopyModule thisWBname, "Module3", schedule_wb
                    CopyModule thisWBname, "TANFMod", schedule_wb
                    CopyForm thisWBname, "MASelectForms", schedule_wb
                    CopyForm thisWBname, "SelectDate", schedule_wb
                    CopyForm thisWBname, "SelectForms", schedule_wb
                    CopyForm thisWBname, "SelectTime", schedule_wb
                ElseIf program = "MA PE" Then
                    CopyModule thisWBname, "CAO_Appointment", schedule_wb
                    CopyModule thisWBname, "CashMemos", schedule_wb
                    CopyModule thisWBname, "Finding_Memo", schedule_wb
                    CopyModule thisWBname, "Module1", schedule_wb
                    CopyModule thisWBname, "Module3", schedule_wb
                    CopyModule thisWBname, "TANFMod", schedule_wb
                    CopyForm thisWBname, "MASelectForms", schedule_wb
                    CopyForm thisWBname, "SelectDate", schedule_wb
                    CopyForm thisWBname, "SelectForms", schedule_wb
                    CopyForm thisWBname, "SelectTime", schedule_wb
                End If
                
                'save workbook
                schedule_wb.Close True
        End If
    End With
    End If
Next i

' Save Display Status Bar
Application.DisplayStatusBar = False

' Keep Screen from updating
Application.ScreenUpdating = True

End Sub
