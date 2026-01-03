Attribute VB_Name = "populate_snap_neg_delimited_mod"
Sub populate_snap_neg_delimited()
Dim wb_sch As Workbook, review_number As String, case_row As Range

Set wb_sch = wb
    For Each ws In wb_sch.Worksheets
        If Val(ws.Name) > 1000 Then
            review_number = ws.Name
            Exit For
        End If
    Next
        
    'Find max row in bis case record
        maxrow_bis_case = wb_bis.Worksheets(1).Cells.Find(What:="*", _
            SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
    'Find max column in bis case record
        LastColumn_bis_case = wb_bis.Worksheets(1).Cells.Find(What:="*", After:=[A1], _
            SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    
    'Find what row in the bis case file the review number is on
    With wb_bis.Worksheets(1).Range("A2:A" & maxrow_bis_case)
        Set case_row = .Find(review_number, LookIn:=xlValues)
    End With
    
    'put todays date in Date Assigned cells in schedule
    wb_sch.Worksheets(review_number).Range("C16") = Application.WorksheetFunction.Text(month(Date), "00")
    wb_sch.Worksheets(review_number).Range("F16") = Application.WorksheetFunction.Text(Day(Date), "00")
    wb_sch.Worksheets(review_number).Range("I16") = Year(Date)
    
    'if case is not found in delimited file, then skip rest of processing in this subroutine
    If case_row Is Nothing Then
        Exit Sub
    End If
    
    'put in Action Date
    wb_sch.Worksheets(review_number).Range("S24") = Mid(wb_bis.Worksheets(1).Range("K" & case_row.Row), 5, 2)
    wb_sch.Worksheets(review_number).Range("V24") = Right(wb_bis.Worksheets(1).Range("K" & case_row.Row), 2)
    wb_sch.Worksheets(review_number).Range("Y24") = Left(wb_bis.Worksheets(1).Range("K" & case_row.Row), 4)

    'put in Notice Date except for suspensions
    If wb_bis.Worksheets(1).Range("C" & case_row.Row) <> "S" Then
        wb_sch.Worksheets(review_number).Range("G24") = Mid(wb_bis.Worksheets(1).Range("S" & case_row.Row), 5, 2)
        wb_sch.Worksheets(review_number).Range("J24") = Right(wb_bis.Worksheets(1).Range("S" & case_row.Row), 2)
        wb_sch.Worksheets(review_number).Range("M24") = Left(wb_bis.Worksheets(1).Range("S" & case_row.Row), 4)
     End If

    'put in type of negativity
    Select Case wb_bis.Worksheets(1).Range("C" & case_row.Row)
        Case "A"
            wb_sch.Worksheets(review_number).Range("AE24") = 1
            action_type = "Denial"
        Case "C"
            wb_sch.Worksheets(review_number).Range("AE24") = 2
            action_type = "Termination"
        Case "S"
            wb_sch.Worksheets(review_number).Range("AE24") = 3
            action_type = "Suspension"
    End Select
    
    'put in first sentence in text box
    wb_sch.Worksheets(review_number).Shapes("Text Box 17").TextFrame.Characters.Text = _
        "The action being reviewed is the SNAP " & _
        action_type & " of " & wb_sch.Worksheets(review_number).Range("S24") & "/" & _
        wb_sch.Worksheets(review_number).Range("V24") & "/" & _
        wb_sch.Worksheets(review_number).Range("Y24") & "."
End Sub
