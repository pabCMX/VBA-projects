Attribute VB_Name = "Drop"
Sub Drop()

Answer = MsgBox("This will delete values on your schedule.  Are you sure you want to delete them?", vbQuestion + vbYesNo, "???")
    

    If Answer = vbNo Then
        'Code for No button Press
        Exit Sub
    End If
 'ActiveSheet.Unprotect "QC"

If Left(ActiveSheet.Name, 1) = "5" Then

    Range("K22:O22") = ""
    Range("Y22:O22") = ""
    'Range("C50:AI50") = "" Updated 10/1/2023 need to fill out for drop
    'Range("E55:AI55") = "" Updated 10/1/2023 need to fill out for drop
    'Range("C62:AI62") = "" Updated 5/3/2024 need to fill out for drop
    'Range("D69:M69") = "" Updated 5/3/2024 need to fill out for drop
    'Range("D76:AJ76") = "" Updated 5/3/2024 need to fill out for drop
    'Range("N82:AH82") = "" Updated 5/3/2024 need to fill out for drop
    'Range("Y89:AN89") = "" Updated 5/3/2024 need to fill out for drop
    'Range("Y92:AN92") = "" Updated 5/3/2024 need to fill out for drop
    'Range("Y95:AN95") = "" Updated 5/3/2024 need to fill out for drop
    'Range("Y98:AN98") = "" Updated 5/3/2024 need to fill out for drop
    'Range("Y101:AN101") = "" Updated 5/3/2024 need to fill out for drop
    'Range("Y104:AN104") = "" Updated 5/3/2024 need to fill out for drop
    'Range("Y107:AN107") = "" Updated 5/3/2024 need to fill out for drop
    'Range("Y110:AN110") = "" Updated 5/3/2024 need to fill out for drop
    'Range("Y113:AN113") = "" Updated 5/3/2024 need to fill out for drop
    'Range("Y116:AN116") = "" Updated 5/3/2024 need to fill out for drop
    'Range("Y119:AN119") = "" Updated 5/3/2024 need to fill out for drop
    'Range("Y122:AN122") = "" Updated 5/3/2024 need to fill out for drop
    Range("A131:AN143") = ""
    'Range("B155:C155") = ""
    'Range("B157:C157") = ""
    'Range("B159:C159") = ""

'Section 2
    For i = 29 To 43 Step 2
        Range("B" & i & ":AN" & i) = ""
    Next

ElseIf Left(ActiveSheet.Name, 1) = "1" Then
    'doesn't clear last 4 cells for resources
    Range("A16:Y16") = ""
    Range("C20:AN20") = ""
    Range("B24:AN24") = ""
    
    'clear row 85
    Range("A85:AT85") = ""
    
    For i = 30 To 44 Step 2
        Range("A" & i & ":AQ" & i) = ""
    Next
    
    For i = 50 To 56 Step 2
        Range("C" & i & ":AO" & i) = ""
    Next
    
    For i = 61 To 67 Step 2
        Range("C" & i & ":AR" & i) = ""
    Next
    
    For i = 72 To 82 Step 2
        Range("F" & i & ":AN" & i) = ""
    Next
    
'Worksheets("TANF Workbook").Shapes("CB 203").OLEFormat.Object.Value = 0
'Worksheets("TANF Workbook").Shapes("CB 198").OLEFormat.Object.Value = 0
'Worksheets("TANF Workbook").Shapes("CB 211").OLEFormat.Object.Value = 0
'Worksheets("TANF Workbook").Shapes("CB 214").OLEFormat.Object.Value = 0
'Worksheets("TANF Workbook").Shapes("CB 396").OLEFormat.Object.Value = 0
'Worksheets("TANF Workbook").Shapes("CB 449").OLEFormat.Object.Value = 0
'Worksheets("TANF Workbook").Shapes("CB 393").OLEFormat.Object.Value = 0
'Worksheets("TANF Workbook").Shapes("OB 207").OLEFormat.Object.Value = 0
'Worksheets("TANF Workbook").Shapes("OB 256").OLEFormat.Object.Value = 0
'Worksheets("TANF Workbook").Shapes("OB 388").OLEFormat.Object.Value = 0

ElseIf Left(ActiveSheet.Name, 1) = "2" Then

    Range("A27:AN27") = ""
    Range("B32:AN32") = ""
    Range("I35:X35") = ""
    Range("I38:X38") = ""
    Range("I41:X41") = ""
    

    For i = 51 To 73 Step 2
        Range("B" & i & ":AQ" & i) = ""
    Next
    
    For j = 78 To 84 Step 2
        Range("B" & i & ":AN" & i) = ""
    Next
    
    For k = 96 To 112 Step 2
        Range("B" & i & ":AQ" & i) = ""
    Next

    
End If

 'ActiveSheet.Protect "QC"

End Sub
