Option Explicit

' Public counters
Public rswr As Long, qcwr As Long, plwr As Long, hiwr As Long, efwr As Long
Public revidval As Long
Public inWB As Workbook, outWB As Workbook, inWS As Worksheet
Public thissht As Worksheet

Sub Find_Write_Database_Files()
    On Error GoTo ErrorHandler

    ' 1. INITIAL SETUP FOR SPEED
    With Application
        .ScreenUpdating = False
        .DisplayStatusBar = True
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With

    Dim i As Long, j As Long, n As Long
    Dim maxrow As Long, maxrowex As Long
    Dim program As String, exname As String, exnumstr As String
    Dim reviewtxt As String, disp_code As Variant
    Dim pathdir As String, BasePath As String, CaseFolderPath As String
    Dim CaseFolderName As String, FinalFilePath As String, FileNameFound As String
    Dim monthstr As String, mName As String, yStr As String
    Dim sPath As String, SrceFile As String, exceloutfile As String, databasename As String
    
    ' ADO Objects
    Dim cnt As Object, rs As Object
    Dim wsname As String
    Dim filenum As Integer
    Dim rowCursor As Long, c As Long
    Dim valToInsert As Variant
    Dim colCountArr As Variant, tableArr As Variant, maxRowsArr As Variant
    Dim headerRange As Variant, dataRange As Variant
    Dim FieldMap() As Integer ' Array to store column-to-field mapping

    Set thissht = ActiveSheet
    sPath = ActiveWorkbook.Path

    ' 2. NETWORK & PATH SETUP
    Dim WshNetwork As Object, oDrives As Object
    Set WshNetwork = CreateObject("WScript.Network")
    Set oDrives = WshNetwork.EnumNetworkDrives

    Dim DLetter As String
    For i = 0 To oDrives.Count - 1 Step 2
        If LCase(oDrives.Item(i + 1)) = "\\hsedcprapfpp001\oim\pwimdaubts04\data\stat" Then
            DLetter = oDrives.Item(i) & "\DQC\": Exit For
        ElseIf LCase(oDrives.Item(i + 1)) = "\\hsedcprapfpp001\oim\pwimdaubts04\data\stat\dqc" Then
            DLetter = oDrives.Item(i) & "\": Exit For
        End If
    Next i

    If DLetter = "" Then Err.Raise 9999, , "Network Drive not found."
    pathdir = DLetter & "Schedules by Examiner Number\"
    If Dir(pathdir, vbDirectory) = "" Then Err.Raise 9999, , "Examiner Directory not found."

    ' 3. LOAD INPUT DATA INTO MEMORY (SPEED OPTIMIZATION)
    maxrow = thissht.Range("E" & thissht.Rows.Count).End(xlUp).Row
    maxrowex = thissht.Range("L" & thissht.Rows.Count).End(xlUp).Row
    
    ' Load Main Data (Cols E to G) and Examiner Lookup (Cols K to L) into Arrays
    Dim vInput As Variant, vExLookup As Variant
    vInput = thissht.Range("E1:G" & maxrow).Value ' E=RevNum, F=Month, G=ExNum
    vExLookup = thissht.Range("K1:L" & maxrowex).Value ' K=Name, L=Number

    ' Prepare Output File
    SrceFile = sPath & "\FO Databases\TANF_Template.xlsx"
    exceloutfile = sPath & "\TANF Database Input " & Format(Date, "mm-dd-yyyy") & ".xlsx"
    If Dir(exceloutfile) <> "" Then Kill exceloutfile
    FileCopy SrceFile, exceloutfile
    Set outWB = Workbooks.Open(exceloutfile)
    
    rswr = 1: qcwr = 1: plwr = 1: hiwr = 1: efwr = 1

    ' 4. MAIN PROCESSING LOOP (MEMORY BASED)
    For i = 2 To maxrow
        Application.StatusBar = "Processing " & i - 1 & "/" & maxrow - 1
        DoEvents

        ' A. Parse Review Number (vInput array is 1-based)
        reviewtxt = CStr(vInput(i, 1)) ' Col E
        If Left(reviewtxt, 1) = "0" Then reviewtxt = Mid(reviewtxt, 2)
        If Left(reviewtxt, 1) <> "1" Then GoTo NextIteration
        
        program = "TANF"

        ' B. Find Examiner Name (Lookup in Array)
        exname = ""
        Dim currentExNum As String
        currentExNum = CStr(vInput(i, 3)) ' Col G
        
        For j = 2 To maxrowex
            If CStr(vExLookup(j, 2)) = currentExNum Then
                exname = vExLookup(j, 1)
                Exit For
            End If
        Next j

        If exname = "" Then GoTo NextIteration

        exnumstr = Format(currentExNum, "00")
        If Left(exnumstr, 1) = "0" Then exnumstr = Right(exnumstr, 1)

        ' C. Parse Date
        monthstr = CStr(vInput(i, 2)) ' Col F
        If Len(monthstr) <> 6 Then GoTo NextIteration
        
        yStr = Left(monthstr, 4)
        mName = MonthName(Val(Right(monthstr, 2)))

        ' D. Dynamic Path Hunting
        BasePath = pathdir & exname & " - " & exnumstr & "\" & program & "\" & _
                   "Review Month " & mName & " " & yStr & "\"

        CaseFolderName = Dir(BasePath & reviewtxt & " - *", vbDirectory)

        If CaseFolderName <> "" And (GetAttr(BasePath & CaseFolderName) And vbDirectory) = vbDirectory Then
            CaseFolderPath = BasePath & CaseFolderName & "\"
            FileNameFound = Dir(CaseFolderPath & "Review Number " & reviewtxt & "*.xls*")
            
            If FileNameFound <> "" Then
                FinalFilePath = CaseFolderPath & FileNameFound
                
                Set inWB = Workbooks.Open(Filename:=FinalFilePath, UpdateLinks:=0, ReadOnly:=True)
                
                On Error Resume Next
                Set inWS = inWB.Sheets(reviewtxt)
                If Err.Number <> 0 Then
                    inWB.Close False
                    Err.Clear
                    GoTo NextIteration
                End If
                On Error GoTo ErrorHandler

                disp_code = inWS.Range("AI10").Value
                revidval = i - 1

                ' Extract Data
                Call revsum(outWB.Sheets("Review_Summary_dtl"))
                If disp_code = 1 Then
                    Call qcinfo(outWB.Sheets("QC_Case_Info_dtl"))
                    Call plinfo(outWB.Sheets("Person_Level_Info_dtl"))
                    Call hhinc(outWB.Sheets("Household_Income_dtl"))
                    Call errfind(outWB.Sheets("Error_Findings_dtl"))
                End If

                inWB.Close False
                Set inWB = Nothing
            End If
        End If

NextIteration:
    Next i

    ' 5. DATABASE TRANSFER (OPTIMIZED MAPPING)
    outWB.Save
    Application.StatusBar = "Transferring to Database..."
    DoEvents

    SrceFile = sPath & "\FO Databases\TANF_Blank.mdb"
    databasename = sPath & "\TANF1 " & Format(Date, "mm-dd-yyyy") & ".mdb"
    filenum = 1
    Do Until Dir(databasename) = ""
        filenum = filenum + 1
        databasename = sPath & "\TANF" & filenum & " " & Format(Date, "mmddyyyy") & ".mdb"
    Loop
    FileCopy SrceFile, databasename

    Set cnt = CreateObject("ADODB.Connection")
    cnt.Open "Provider=Microsoft.Ace.OLEDB.12.0;Data Source=" & databasename & ";Persist Security Info=False;"
    
    tableArr = Array("Review_Summary_dtl", "QC_Case_Info_dtl", "Person_Level_Info_dtl", "Household_Income_dtl", "Error_Findings_dtl")
    maxRowsArr = Array(rswr, qcwr, plwr, hiwr, efwr)
    colCountArr = Array(20, 30, 25, 10, 20) ' Max columns to scan

    For n = LBound(tableArr) To UBound(tableArr)
        If maxRowsArr(n) > 1 Then
            wsname = tableArr(n)
            
            Dim wsOut As Worksheet
            Set wsOut = outWB.Sheets(wsname)
            
            Set rs = CreateObject("ADODB.Recordset")
            rs.Open wsname, cnt, 1, 3, 2
            
            ' A. Read Excel Data to Memory
            Dim lastCol As Long
            lastCol = colCountArr(n)
            headerRange = wsOut.Range(wsOut.Cells(1, 1), wsOut.Cells(1, lastCol)).Value
            dataRange = wsOut.Range(wsOut.Cells(2, 1), wsOut.Cells(maxRowsArr(n), lastCol)).Value
            
            ' B. Pre-Calculate Column Mapping (Optimization)
            ReDim FieldMap(1 To lastCol)
            Dim accessField As Object, hName As String
            
            For c = 1 To lastCol
                hName = Trim(CStr(headerRange(1, c)))
                FieldMap(c) = -1 ' Default to not found
                
                If hName <> "" Then
                    ' Normalize Header
                    Dim normH As String: normH = Replace(LCase(hName), " ", "_")
                    
                    ' Find Field Index
                    For j = 0 To rs.Fields.Count - 1
                        If LCase(rs.Fields(j).Name) = normH Or LCase(rs.Fields(j).Name) = LCase(hName) Then
                            FieldMap(c) = j
                            Exit For
                        End If
                    Next j
                End If
            Next c
            
            ' C. Fast Loop
            For rowCursor = 1 To UBound(dataRange, 1)
                rs.AddNew
                For c = 1 To UBound(dataRange, 2)
                    ' Only process if we have a valid field map and data exists
                    If FieldMap(c) > -1 Then
                        valToInsert = dataRange(rowCursor, c)
                        If Not IsError(valToInsert) Then
                            If Len(CStr(valToInsert)) > 0 Then
                                On Error Resume Next
                                rs.Fields(FieldMap(c)).Value = valToInsert
                                If Err.Number <> 0 Then
                                    MsgBox "Write Error!" & vbCrLf & _
                                           "Table: " & wsname & vbCrLf & _
                                           "Row: " & rowCursor & vbCrLf & _
                                           "Col: " & headerRange(1, c) & vbCrLf & _
                                           "Val: " & valToInsert & vbCrLf & _
                                           Err.Description, vbCritical
                                    cnt.Close: outWB.Close: Exit Sub
                                End If
                                On Error GoTo ErrorHandler
                            End If
                        End If
                    End If
                Next c
                rs.Update
            Next rowCursor
            
            rs.Close: Set rs = Nothing
        End If
    Next n

    cnt.Close: Set cnt = Nothing
    outWB.Close True: Set outWB = Nothing
    If Dir(exceloutfile) <> "" Then Kill exceloutfile

    MsgBox "Complete!", vbInformation

CleanExit:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

ErrorHandler:
    MsgBox "Critical Error: " & Err.Description & " (" & Err.Number & ")", vbCritical
    If Not inWB Is Nothing Then inWB.Close False
    If Not outWB Is Nothing Then outWB.Close True
    Resume CleanExit
End Sub

' -----------------------------------------------------------
' LIGHTWEIGHT HELPER SUBS (Direct Write)
' -----------------------------------------------------------

Sub revsum(ws As Worksheet)
    rswr = rswr + 1
    ws.Cells(rswr, 1).Value = revidval
    ws.Cells(rswr, 2).Value = inWS.Range("A10").Value
    ws.Cells(rswr, 3).Value = inWS.Range("I10").Value
    ws.Cells(rswr, 4).Value = inWS.Range("Q10").Value
    ws.Cells(rswr, 5).Value = inWS.Range("S10").Value
    
    Dim d As Variant: d = inWS.Range("AB10").Value
    If IsNumeric(d) And Len(CStr(d)) >= 6 Then ws.Cells(rswr, 6).Value = DateSerial(Val(Right(d, 4)), Val(Left(d, 2)), 1)

    If inWS.Range("AI10").Value = 1 Then
        If IsNumeric(inWS.Range("AO10").Value) Then ws.Cells(rswr, 7).Value = Round(Val(inWS.Range("AO10").Value) + 0.001, 0)
    End If

    ws.Cells(rswr, 8).Value = Cln(inWS.Range("AL10").Value, "B")
    ws.Cells(rswr, 9).Value = Cln(inWS.Range("Y10").Value, "BB")
    ws.Cells(rswr, 10).Value = inWS.Range("U10").Value
    ws.Cells(rswr, 11).Value = inWS.Range("AI10").Value

    On Error Resume Next
    ws.Cells(rswr, 12).Value = inWB.Sheets("TANF Workbook").Range("G33").Value
    On Error GoTo 0
    
    ws.Cells(rswr, 13).Value = inWS.Range("AO3").Value & inWS.Range("AP3").Value
    ws.Cells(rswr, 15).Value = Cln(inWS.Range("AB85").Value, "B")
End Sub

Sub qcinfo(ws As Worksheet)
    qcwr = qcwr + 1
    ws.Cells(qcwr, 1).Value = revidval
    ws.Cells(qcwr, 2).Value = inWS.Range("V20").Value
    ws.Cells(qcwr, 3).Value = inWS.Range("Y20").Value
    ws.Cells(qcwr, 4).Value = inWS.Range("J16").Value
    ws.Cells(qcwr, 5).Value = inWS.Range("O20").Value
    ws.Cells(qcwr, 6).Value = inWS.Range("U16").Value
    ws.Cells(qcwr, 7).Value = inWS.Range("A16").Value
    ws.Cells(qcwr, 8).Value = inWS.Range("L16").Value
    ws.Cells(qcwr, 9).Value = Val(inWS.Range("W16").Value)
    ws.Cells(qcwr, 10).Value = inWS.Range("Z16").Value
    ws.Cells(qcwr, 11).Value = inWS.Range("AE16").Value
    ws.Cells(qcwr, 12).Value = inWS.Range("AJ16").Value
    ws.Cells(qcwr, 13).Value = inWS.Range("AO16").Value
    ws.Cells(qcwr, 14).Value = inWS.Range("C20").Value
    ws.Cells(qcwr, 15).Value = inWS.Range("I20").Value
    ws.Cells(qcwr, 16).Value = inWS.Range("Q20").Value
    ws.Cells(qcwr, 17).Value = inWS.Range("AB20").Value
    ws.Cells(qcwr, 18).Value = inWS.Range("AH20").Value
    ws.Cells(qcwr, 19).Value = inWS.Range("AN20").Value
    ws.Cells(qcwr, 20).Value = inWS.Range("B24").Value
    ws.Cells(qcwr, 21).Value = inWS.Range("U24").Value
    ws.Cells(qcwr, 22).Value = inWS.Range("N20").Value
    ws.Cells(qcwr, 23).Value = inWS.Range("AN24").Value
End Sub

Sub plinfo(ws As Worksheet)
    Dim j As Long, ln As Long
    For j = 30 To 44 Step 2
        If inWS.Range("A" & j).Value = "" Then Exit For
        plwr = plwr + 1: ln = ln + 1
        
        ws.Cells(plwr, 1).Value = revidval
        ws.Cells(plwr, 2).Value = inWS.Range("A" & j).Value
        ws.Cells(plwr, 3).Value = inWS.Range("D" & j).Value
        ws.Cells(plwr, 4).Value = inWS.Range("E" & j).Value
        ws.Cells(plwr, 5).Value = inWS.Range("G" & j).Value
        ws.Cells(plwr, 6).Value = inWS.Range("J" & j).Value
        ws.Cells(plwr, 7).Value = Val(inWS.Range("M" & j).Value)
        ws.Cells(plwr, 8).Value = inWS.Range("P" & j).Value
        ws.Cells(plwr, 9).Value = inWS.Range("R" & j).Value
        ws.Cells(plwr, 10).Value = inWS.Range("T" & j).Value
        ws.Cells(plwr, 11).Value = inWS.Range("V" & j).Value
        ws.Cells(plwr, 12).Value = inWS.Range("Y" & j).Value
        ws.Cells(plwr, 13).Value = inWS.Range("AC" & j).Value
        ws.Cells(plwr, 14).Value = inWS.Range("AF" & j).Value
        ws.Cells(plwr, 15).Value = inWS.Range("AK" & j).Value
        ws.Cells(plwr, 16).Value = inWS.Range("AN" & j).Value
        ws.Cells(plwr, 17).Value = inWS.Range("AQ" & j).Value
        ws.Cells(plwr, 18).Value = ln
    Next j
End Sub

Sub hhinc(ws As Worksheet)
    Dim j As Long, k As Long
    For j = 50 To 56 Step 2
        If inWS.Range("C" & j).Value = "" Then Exit For
        For k = 7 To 37 Step 10
            If inWS.Cells(j, k).Value = "" Then Exit For
            hiwr = hiwr + 1
            ws.Cells(hiwr, 1).Value = revidval
            ws.Cells(hiwr, 2).Value = inWS.Cells(j, k + 4).Value
            ws.Cells(hiwr, 4).Value = inWS.Range("C" & j).Value
            ws.Cells(hiwr, 5).Value = inWS.Cells(j, k).Value
        Next k
    Next j
End Sub

Sub errfind(ws As Worksheet)
    Dim j As Long, ln As Long
    For j = 61 To 67 Step 2
        If inWS.Range("F" & j).Value = "" Then Exit For
        efwr = efwr + 1: ln = ln + 1
        
        ws.Cells(efwr, 1).Value = revidval
        ws.Cells(efwr, 2).Value = inWS.Range("F" & j).Value
        ws.Cells(efwr, 3).Value = inWS.Range("AR" & j).Value
        ws.Cells(efwr, 4).Value = inWS.Range("AD" & j).Value
        ws.Cells(efwr, 5).Value = inWS.Range("AH" & j).Value
        ws.Cells(efwr, 6).Value = inWS.Range("T" & j).Value
        ws.Cells(efwr, 7).Value = inWS.Range("C" & j).Value
        ws.Cells(efwr, 8).Value = inWS.Range("X" & j).Value
        
        If IsDate(inWS.Range("AL" & j).Value) Then
            Dim d As Date: d = inWS.Range("AL" & j).Value
            ws.Cells(efwr, 9).Value = DateSerial(Year(d), Month(d), 1)
        End If
        
        ws.Cells(efwr, 10).Value = inWS.Range("O" & j).Value
        ws.Cells(efwr, 11).Value = inWS.Range("J" & j).Value
        ws.Cells(efwr, 12).Value = ln
    Next j
End Sub

' Function to safely clean strings
Function Cln(v As Variant, d As String) As String
    If IsError(v) Then Cln = d: Exit Function
    If Trim(CStr(v)) = "" Or InStr(CStr(v), "-") > 0 Then Cln = d Else Cln = CStr(v)
End Function