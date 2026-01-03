VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Workbook_BeforePrint(Cancel As Boolean)

Dim thisws As Worksheet, thissum As Long

    ' find name of schedule spreadsheet ~ SNAP POSITIVE
    For Each ws In ThisWorkbook.Worksheets
    If Left(ws.Name, 2) = "50" Or Left(ws.Name, 2) = "51" Or Left(ws.Name, 2) = "55" Then
        Set thisws = ws
        program = "SNAP"
    ElseIf Left(ws.Name, 2) = "20" Or Left(ws.Name, 2) = "21" Or Left(ws.Name, 2) = "23" Then
        Set thisws = ws
        program = "MA"
    End If
    Next ws
    
If program = "SNAP" Then
If thisws.Name = ActiveSheet.Name Then

If thisws.Range("C22") = "1" Then
    If thisws.Range("AJ76") = "" Then
        MsgBox ("Please enter a value into item 42 (Homeless).")
Cancel = True
    End If
End If

'Make sure length of cert period is greater than 0 when disposition = 1

If thisws.Range("C22") = "1" Then
    If thisws.Range("T50") = "0" Or thisws.Range("T50") = "" Or thisws.Range("T50") = "00" Then
        MsgBox ("Please enter a value greater than 0 into item 22 (length of cert. period).")
        Range("T50") = ""
Cancel = True
    End If
End If

If thisws.Range("C22") = "1" Then
    If thisws.Range("S55") = "" Then
        MsgBox ("Please enter a value into item 27 (Authorized Representative).")
Cancel = True
    End If
End If

'if box 1 of item 44 (use of SUA) = 1 then Item 45 (utlities) should = 0

If thisws.Range("C22") = "1" Or thisws.Range("C22") = "4" Then
    If thisws.Range("AH82") = 0 Then
        thisws.Range("W82") = 1
        thisws.Range("AA82") = "-"
    End If
End If

If thisws.Range("C22") = "1" Or thisws.Range("C22") = "4" Then
    If thisws.Range("AH82") <> 0 And thisws.Range("W82") = 1 Then
        MsgBox ("Box 1 of Item 44 (Use of SUA) cannot = 1, if Item 45 (Utilites) is greater than 0")
    Range("W82") = ""
    Range("AA82") = ""
Cancel = True
    'End If
'End If

'If thisws.Range("C22") = "1" Or thisws.Range("C22") = "4" Then
    ElseIf thisws.Range("W82") <> 1 And (thisws.Range("AA82") = "-" Or thisws.Range("AA82") = "") Then
        MsgBox ("If Box 1 of Item 44 (Usage of SUA) is not = 1, then Box 2 of Item 44 (Proration of SUA) must be a 1 or 2.")
        Range("AA82") = ""
Cancel = True
    End If
End If

If thisws.Range("C22") = "1" Then
    If thisws.Range("AB50") = "" Then
        MsgBox ("Please enter a value into item 23 (Allotment Adjustment).")
Cancel = True
    End If
End If

If thisws.Range("K22") = 1 Or thisws.Range("K22") = 2 Or thisws.Range("K22") = 3 Then
    thissum = 0
    For i = 89 To 122 Step 3
        If thisws.Range("E" & i) = 1 Then
            thissum = thissum + thisws.Range("AJ" & i)
        End If
    Next i
    If thissum < thisws.Range("O76") - 5 Then
        MsgBox "Sum of Item 58 (Dependent Care Cost) which equals $" & thissum & " must be greater than or $5 less than Item 39 (Dependent Care Deduction) which equals $" & thisws.Range("O76") & "."
        Cancel = True
    End If
ElseIf thisws.Range("K22") = "4" Then
    If thisws.Range("B89") <> "" Or thisws.Range("B131") <> "" Then
        Ineligible = MsgBox("This schedule is an ineligible review.  Please click Yes to clear sections 4 and 5.", vbYesNo)
            If Ineligible = vbYes Then
                Range("A89:AN122") = ""
                Range("A131:AN143") = ""
                Cancel = False
            Else
                Cancel = True
            End If

    End If
End If

'Make sure there is an IEVS code when the case is an error case

'If thisws.Range("Y22") > "0" Then
'    If thisws.Range("C155") = "" Then
'        MsgBox ("Please enter an IEVS code on line 2 block 2 of the supplemental section.")
'Cancel = True
'    End If
'End If

End If


 
' find name of schedule spreadsheet ~ MEDICAL ASSISTANCE POSITIVE
ElseIf program = "MA" Then
If thisws.Name = ActiveSheet.Name Then

'enter a citizenship and identity code
'If thisws.Range("F16") = "1" Then
'    If thisws.Range("O118") = "" Then
'        MsgBox ("Please enter a Citizenship and Identity code in the supplemental section.")
'Cancel = True
'    End If
'End If

'enter a voter id
'If thisws.Range("F16") = "1" Then
 '   If thisws.Range("T118") = "" Then
 '       MsgBox ("Please enter a Voter ID code in the supplemental section.")
'Cancel = True
'    End If
'End If

'enter a IEVS type
'If thisws.Range("D96") <> "" Then
'    If thisws.Range("Y118") = "" Then
'        MsgBox ("Please enter an IEVS code in the supplemental section.")
'Cancel = True
'    End If
'End If

'enter a renewal type
'If thisws.Range("F16") = "1" Then
'    If thisws.Range("AE118") = "" Then
'        MsgBox ("Please enter a Renewal Type code in the supplemental section.")
'Cancel = True
'    End If
'End If



End If

End If

End Sub



