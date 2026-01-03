Attribute VB_Name = "TransPopulate"

Sub Transmittal()

'
' addresses Macro
' Macro recorded 3/25/2008 by vamiller

Dim program As String
Dim month As Long
Dim Start_Range As Integer
Dim End_Range As Integer
Dim monthstr As String
Dim fileName As String
Dim TempStr As String

' Find file of records file to open
    'ChDrive "\\dhs\share\oim\pwimdaubts04\data"
    'ChDir "\stat\DQC"
    input_file = Application _
    .GetOpenFilename(" File of Records(*.xlsx), *.xlsx", , _
    "Select File of Record file")
    If input_file = "False" Then Exit Sub
    
    Application.ScreenUpdating = False

'Read which program and month from user formpop1 input
    program = Cells(7, 23)
    month = Cells(7, 26)

'create new file with program and month as name
    monthstr = Application.WorksheetFunction.Text(month, "MMMM YYYY")
    fileName = "Transmittals for " & program & " " & monthstr & ".xlsx"

    ' create new workbook
    Application.DisplayAlerts = False
    Workbooks.Add
          ActiveWorkbook.SaveAs fileName:= _
          fileName, FileFormat:=xlNormal, _
          Password:="", WriteResPassword:="", ReadOnlyRecommended:=False _
          , CreateBackup:=False
    Application.DisplayAlerts = True
    
   'Open file of records
    Workbooks.Open fileName:=input_file, UpdateLinks:=False
    
    ' copy worksheet from file of records
    If program = "TANF" Or program = "GA" Or program = "FS Supplemental" Or program = "FS Positive" Or program = "FS Negative" Then
        Sheets("FS Cash main file").Select
        Sheets("FS Cash main file").Copy before:=Workbooks(fileName).Sheets(2)
        Sheets("FS Cash main file").Name = "Temp"
    ElseIf program = "MA Negative" Or program = "MA Positive" Then
        Sheets("MA main file").Select
        Sheets("MA main file").Copy before:=Workbooks(fileName).Sheets(2)
        Sheets("MA main file").Name = "Temp"
    Else
        Sheets("CAR main file").Select
        Sheets("CAR main file").Copy before:=Workbooks(fileName).Sheets(2)
        Sheets("CAR main file").Name = "Temp"
    End If
   
    ' Find input file name
    testarr = Split(input_file, "\")
    upper = UBound(testarr)
    file_name = testarr(upper)
    
    Application.DisplayAlerts = False
        
     'Close file of records
     Windows(file_name).Close
     
    Application.DisplayAlerts = True
   
   'Activate Transmittal spreadsheet
         Windows(fileName).Activate
         Sheets("Temp").Select
    
    ' Find maximum row in temp spreadsheet
    maxrow = ActiveSheet.Cells.Find(What:="*", _
        SearchDirection:=xlPrevious, _
        SearchOrder:=xlByRows).Row
     
    'test for month in file
    flag = 0
    For i = 2 To maxrow
       If Cells(i, 2) = month Then
          flag = 1
          Exit For
      End If
    Next i

    ' If month not found, display message and exit
    If flag = 0 Then
      MsgBox month & " Month not Found. Check File of Records."
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
    Case "MA Positive"
         Start_Range = 20
         End_Range = 23
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
    Case "MA Negative"
         Start_Range = 80
         End_Range = 82
    End Select
    
'delete unwanted cases and months
    For i = maxrow To 1 Step -1
      If Val(Left(Cells(i, 1), 2)) < Start_Range Or Val(Left(Cells(i, 1), 2)) > End_Range Or Cells(i, 2) <> month Then
       Rows(i).Delete
     End If
    Next i

'determining if there are any schedules in temp file
    If Range("A1") = "" Then
        MsgBox ("No schedules found for selected program and month")
        Windows("populate.xlsm").Activate
        Sheets("Populate").Select
        Workbooks(fileName).Close (False)
        End
    End If
      
     ' Find maximum row in temp spreadsheet
    maxrow = ActiveSheet.Cells.Find(What:="*", _
        SearchDirection:=xlPrevious, _
        SearchOrder:=xlByRows).Row
 
    ' Switch to populate workbook and copy transmittal worksheet
    Windows("populate.xlsm").Activate
    
    Sheets("Transmittal").Select
    Sheets("Transmittal").Copy before:=Workbooks(fileName).Sheets(1)
         
    Application.DisplayAlerts = False
  
  'deleting extra sheets
    Sheets("Sheet1").Delete
    Sheets("Sheet2").Delete
    Sheets("Sheet3").Delete
    
    Application.DisplayAlerts = True
    
        'rename sheet by review number
        Sheets("Transmittal").Name = Sheets("Temp").Range("A1")

        'loop through each review number creating transmittal sheets
         For i = 1 To maxrow
         'strtemp = "Processing Form " & i & " of " & maxrow & ". Please be patient..."
         'Application.StatusBar = strtemp
         
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
            If i > 1 Then
                TempStr = Sheets("Temp").Range("A" & i - 1)
                Sheets(TempStr).Select
                Sheets(TempStr).Copy before:=Sheets("Temp")
                TempStr = TempStr & " (2)"
                Sheets(TempStr).Name = Sheets("Temp").Range("A" & i)
            End If
            
            'putting county name
            TempStr = "=VLOOKUP(Temp!R" & i & "C4,[populate.xlsm]Populate!R2C30:R68C31,2,FALSE)"
            Range("C6").FormulaR1C1 = TempStr
            Range("C6").Select
            Selection.Copy
            Range("C6").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
              :=False, Transpose:=False
            
            Range("C6") = Application.WorksheetFunction.Text(Sheets("Temp").Range("D" & i), "00") & " - " & Range("C6") & " CAO"
                      
            Range("C7") = ""
          'Look up District name based on county number
          Select Case Sheets("Temp").Range("D" & i)
            Case 2
                TempStr2 = _
                "=VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R2C15:R9C18,4,FALSE)"
                Range("C7").FormulaR1C1 = TempStr2
             Case 9
                TempStr2 = _
                "=VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R10C15:R11C18,4,FALSE)"
                Range("C7").FormulaR1C1 = TempStr2
            Case 23
                TempStr2 = _
                "=VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R12C15:R13C18,4,FALSE)"
                Range("C7").FormulaR1C1 = TempStr2
            Case 40
                TempStr2 = _
                "=VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R14C15:R15C18,4,FALSE)"
                Range("C7").FormulaR1C1 = TempStr2
            Case 46
                TempStr2 = _
                "=VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R16C15:R17C18,4,FALSE)"
                Range("C7").FormulaR1C1 = TempStr2
            Case 51
                TempStr2 = _
                "=VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R18C15:R36C18,4,FALSE)"
                Range("C7").FormulaR1C1 = TempStr2
            Case 63
                TempStr2 = _
                "=VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R37C15:R38C18,4,FALSE)"
                Range("C7").FormulaR1C1 = TempStr2
            Case 65
                TempStr2 = _
                "=VLOOKUP(Temp!R" & i & "C5,[populate.xlsm]Populate!R39C15:R40C18,4,FALSE)"
                Range("C7").FormulaR1C1 = TempStr2
          End Select
          
            ' If there is a district, then concatenate county and district names
            If Range("C7") <> "" Then
                'Copy values
                Range("C7").Select
                Selection.Copy
                Range("C7").Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
                Range("C7") = Range("C7") & " District"
            
                'concatenate county and district names
                Range("C6") = Range("C6") & " , " & Range("C7")
                Range("C7") = ""
            End If
            
            'Putting First and Last Name
            Range("B10") = Sheets("Temp").Range("I" & i) & " " _
                & Sheets("Temp").Range("H" & i)

            'Putting Case and Review Number
            Range("G10") = Sheets("Temp").Range("F" & i) & " / " & _
                Sheets("Temp").Range("A" & i)
            
            'putting either clerk name or supervisor name
            If program = "MA Positive" Or program = "MA Negative" Then
                'Range("I16") = "George Partilla"
                Range("I17") = "Clerk"
            Else
                'Range("I16") = "Martie Hockenberry"
                Range("I17") = "Clerk"
            End If
            
            Range("A2").Select
        Next i
        
        'Display main sheet in populate
        Windows("populate.xlsm").Activate
        Sheets("Populate").Select
        
        'Delete Temp File
        Application.DisplayAlerts = False
        Workbooks(fileName).Activate
        Sheets("Temp").Delete
        Application.DisplayAlerts = True
        
        Workbooks(fileName).SAVE
        
    End Sub



