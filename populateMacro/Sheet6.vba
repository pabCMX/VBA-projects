VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Private Sub Worksheet_Change(ByVal Target As Range)

' This is the range of cells of person number
  'Const sNAMECELL As String = "B51,B53,B55,B57,B59,B61,B63,B65,B67,B69,B71,B73"
  'Dim allcellsblank As Boolean
  'Dim Rngx As Range
  
' Application.EnableEvents = True   'should be part of Change macro
 
'With Target

' Detect if change is in range of cells of interest
    'If Not Intersect(.Cells, Range(sNAMECELL)) Is Nothing Then
    

' Determine if all cells are blank
        'allcellsblank = True
        'For Each Rngx In Target
            'allcellsblank = allcellsblank And IsEmpty(Rngx)
       'Next
    
' If they are all blank, then blank E&T cells
        'If allcellsblank Then
           ' Cells(.Row, 35) = ""
' If not all blank, then put dashes in E&T cells
        'Else
            'Cells(.Row, 35) = "--"
        'End If
'Putting QC Coverage Code from Schedule to PA 721 sheet
    'ElseIf Not Intersect(.Cells, Range("N16")) Is Nothing Then
    'If Not Intersect(.Cells, Range("N16")) Is Nothing Then
    '    Sheets("PA721").Range("B31") = Left(Range("N16"), 1)
    '    Sheets("PA721").Range("C31") = Right(Range("N16"), 1)
'Putting a check box
    'ElseIf Not Intersect(.Cells, Range("F16")) Is Nothing Then
'Turning Checkboxes off for Completed and Dropped
    'Sheets("PA721").CheckBoxes(2).value = xlOff
    'Sheets("PA721").CheckBoxes(3).value = xlOff
    'If Range("F16") = "1" Then
    '    Sheets("PA721").CheckBoxes(2).value = xlOn
    'ElseIf Range("F16") <> "" Then
     '   Sheets("PA721").CheckBoxes(3).value = xlOn
    'End If
'Putting Initial Eligiblity status from Schedule to PA 721 sheet
    'ElseIf Not Intersect(.Cells, Range("S16")) Is Nothing Then
    '    Sheets("PA721").Range("G37") = Left(Range("S16"), 1)
    '    Sheets("PA721").Range("H37") = Right(Range("S16"), 1)
'Putting Initial Case Liability Error from Schedule to PA 721 sheet
    'ElseIf Not Intersect(.Cells, Range("X16")) Is Nothing Then
    '    TempStr = WorksheetFunction.Text(Range("X16"), "#00000")
    '    Sheets("PA721").Range("C44") = Left(TempStr, 1)
    '    Sheets("PA721").Range("D44") = Mid(TempStr, 2, 1)
    '    Sheets("PA721").Range("E44") = Mid(TempStr, 3, 1)
    '    Sheets("PA721").Range("F44") = Mid(TempStr, 4, 1)
    '    Sheets("PA721").Range("G44") = Right(TempStr, 1)
'Putting Excess Resources from Schedule to PA 721 sheet
    'ElseIf Not Intersect(.Cells, Range("AF16")) Is Nothing Then
    '    TempStr = WorksheetFunction.Text(Range("AF16"), "#00000")
    '    Sheets("PA721").Range("Q44") = Left(TempStr, 1)
    '    Sheets("PA721").Range("R44") = Mid(TempStr, 2, 1)
    '    Sheets("PA721").Range("S44") = Mid(TempStr, 3, 1)
    '    Sheets("PA721").Range("T44") = Mid(TempStr, 4, 1)
    '    Sheets("PA721").Range("U44") = Right(TempStr, 1)
'Putting Error Finding from Schedule to PA 721 sheet
    'ElseIf Not Intersect(.Cells, Range("D96")) Is Nothing Then
    '    Sheets("PA721").Range("B50") = Left(Range("D96"), 1)
    '    Sheets("PA721").Range("C50") = Right(Range("D96"), 1)
    'ElseIf Not Intersect(.Cells, Range("D98")) Is Nothing Then
    '    Sheets("PA721").Range("B53") = Left(Range("D98"), 1)
     '   Sheets("PA721").Range("C53") = Right(Range("D98"), 1)
    'ElseIf Not Intersect(.Cells, Range("D100")) Is Nothing Then
    '    Sheets("PA721").Range("B56") = Left(Range("D100"), 1)
    '    Sheets("PA721").Range("C56") = Right(Range("D100"), 1)
    'ElseIf Not Intersect(.Cells, Range("D102")) Is Nothing Then
    '    Sheets("PA721").Range("B59") = Left(Range("D102"), 1)
    '    Sheets("PA721").Range("C59") = Right(Range("D102"), 1)
    'ElseIf Not Intersect(.Cells, Range("D104")) Is Nothing Then
    '    Sheets("PA721").Range("B62") = Left(Range("D104"), 1)
    '    Sheets("PA721").Range("C62") = Right(Range("D104"), 1)
        
'Putting Case Members with Error from Schedule to PA 721 sheet

    'ElseIf Not Intersect(.Cells, Range("K96")) Is Nothing Then
    '    Sheets("PA721").Range("F50") = Left(Range("K96"), 1)
    '    Sheets("PA721").Range("G50") = Right(Range("K96"), 1)
    'ElseIf Not Intersect(.Cells, Range("K98")) Is Nothing Then
    '    Sheets("PA721").Range("F53") = Left(Range("K98"), 1)
    '    Sheets("PA721").Range("G53") = Right(Range("K98"), 1)
    'ElseIf Not Intersect(.Cells, Range("K100")) Is Nothing Then
    '    Sheets("PA721").Range("F56") = Left(Range("K100"), 1)
    '    Sheets("PA721").Range("G56") = Right(Range("K100"), 1)
    'ElseIf Not Intersect(.Cells, Range("K102")) Is Nothing Then
    '    Sheets("PA721").Range("F59") = Left(Range("K102"), 1)
    '    Sheets("PA721").Range("G59") = Right(Range("K102"), 1)
    'ElseIf Not Intersect(.Cells, Range("K104")) Is Nothing Then
    '    Sheets("PA721").Range("F62") = Left(Range("K104"), 1)
    '    Sheets("PA721").Range("G62") = Right(Range("K104"), 1)
        
    'Bringing over Element Code from Schedule to PA 721
     
    'ElseIf Not Intersect(.Cells, Range("O96")) Is Nothing Then
    '    Sheets("PA721").Range("AC50") = Left(Range("O96"), 1)
    '    Sheets("PA721").Range("AD50") = Mid(Range("O96"), 2, 1)
    '    Sheets("PA721").Range("AE50") = Right(Range("O96"), 1)
    'ElseIf Not Intersect(.Cells, Range("O98")) Is Nothing Then
    '    Sheets("PA721").Range("AC53") = Left(Range("O98"), 1)
    '    Sheets("PA721").Range("AD53") = Mid(Range("O98"), 2, 1)
    '    Sheets("PA721").Range("AE53") = Right(Range("O98"), 1)
    'ElseIf Not Intersect(.Cells, Range("O100")) Is Nothing Then
    '    Sheets("PA721").Range("AC56") = Left(Range("O100"), 1)
    '    Sheets("PA721").Range("AD56") = Mid(Range("O100"), 2, 1)
    '    Sheets("PA721").Range("AE56") = Right(Range("O100"), 1)
    'ElseIf Not Intersect(.Cells, Range("O102")) Is Nothing Then
    '    Sheets("PA721").Range("AC59") = Left(Range("O102"), 1)
    '    Sheets("PA721").Range("AD59") = Mid(Range("O102"), 2, 1)
    '    Sheets("PA721").Range("AE59") = Right(Range("O102"), 1)
    'ElseIf Not Intersect(.Cells, Range("O104")) Is Nothing Then
    '    Sheets("PA721").Range("AC62") = Left(Range("O104"), 1)
    '    Sheets("PA721").Range("AD62") = Mid(Range("O104"), 2, 1)
    '    Sheets("PA721").Range("AE62") = Right(Range("O104"), 1)
   
   'Bringing over Nature Code from Schedule to PA 721
    
    'ElseIf Not Intersect(.Cells, Range("T96")) Is Nothing Then
    '    Sheets("PA721").Range("AH50") = Left(Range("T96"), 1)
    '    Sheets("PA721").Range("AI50") = Mid(Range("T96"), 2, 1)
    '    Sheets("PA721").Range("AJ50") = Right(Range("T96"), 1)
    'ElseIf Not Intersect(.Cells, Range("T98")) Is Nothing Then
    '    Sheets("PA721").Range("AH53") = Left(Range("T98"), 1)
    '    Sheets("PA721").Range("AI53") = Mid(Range("T98"), 2, 1)
    '    Sheets("PA721").Range("AJ53") = Right(Range("T98"), 1)
    'ElseIf Not Intersect(.Cells, Range("T100")) Is Nothing Then
    '    Sheets("PA721").Range("AH56") = Left(Range("T100"), 1)
    '    Sheets("PA721").Range("AI56") = Mid(Range("T100"), 2, 1)
    '    Sheets("PA721").Range("AJ56") = Right(Range("T100"), 1)
    'ElseIf Not Intersect(.Cells, Range("T102")) Is Nothing Then
    '    Sheets("PA721").Range("AH59") = Left(Range("T102"), 1)
    '    Sheets("PA721").Range("AI59") = Mid(Range("T102"), 2, 1)
    '    Sheets("PA721").Range("AJ59") = Right(Range("T102"), 1)
    'ElseIf Not Intersect(.Cells, Range("T104")) Is Nothing Then
    '    Sheets("PA721").Range("AH62") = Left(Range("T104"), 1)
    '    Sheets("PA721").Range("AI62") = Mid(Range("T104"), 2, 1)
    '    Sheets("PA721").Range("AJ62") = Right(Range("T104"), 1)
    
    'Bringing over Agency or Client Code from Schedule to PA 721
    
    'ElseIf Not Intersect(.Cells, Range("X96")) Is Nothing Then
    '    Sheets("PA721").Range("AM50") = Left(Range("X96"), 1)
    '    Sheets("PA721").Range("AN50") = Right(Range("X96"), 1)
    'ElseIf Not Intersect(.Cells, Range("X98")) Is Nothing Then
    '    Sheets("PA721").Range("AM53") = Left(Range("X98"), 1)
    '    Sheets("PA721").Range("AN53") = Right(Range("X98"), 1)
    'ElseIf Not Intersect(.Cells, Range("X100")) Is Nothing Then
    '    Sheets("PA721").Range("AM56") = Left(Range("X100"), 1)
    '    Sheets("PA721").Range("AN56") = Right(Range("X100"), 1)
    'ElseIf Not Intersect(.Cells, Range("X102")) Is Nothing Then
    '    Sheets("PA721").Range("AM59") = Left(Range("X102"), 1)
    '    Sheets("PA721").Range("AN59") = Right(Range("X102"), 1)
    'ElseIf Not Intersect(.Cells, Range("X104")) Is Nothing Then
    '    Sheets("PA721").Range("AM62") = Left(Range("X104"), 1)
    '    Sheets("PA721").Range("AN62") = Right(Range("X104"), 1)
   
    'End If
'End With
 
'Application.EnableEvents = True   'should be part of Change macro

'End Sub

