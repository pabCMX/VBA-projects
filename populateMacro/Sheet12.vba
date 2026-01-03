VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Private Sub Worksheet_Change(ByVal Target As Range)
  
 
  'Application.EnableEvents = True   'should be part of Change macro
  
  ' AMR values
'If (Target.Row = 85) And Target.Column = 16 And Range("AI10") = "1" Then
'    MsgBox "If disposition = 1, Then AMR information should be filled out on the TANF Workbook, not the schedule"
'    Application.EnableEvents = False
'    Range("P85") = ""
'End If
 
' Application.EnableEvents = True   'should be part of Change macro
 
 
'End Sub

Private Sub Worksheet_Change(ByVal Target As Range)

  
 Application.EnableEvents = True   'should be part of Change macro
    Dim TempStr As String, index As Integer
    Dim ws As Worksheet, thisws As Worksheet
    Dim linnum As String, inctype As String

    ' find name of schedule spreadsheet
    For Each ws In ThisWorkbook.Worksheets
    If Left(ws.Name, 2) = "14" Then
        Set thisws = ws
    End If
    Next ws
With Target

'Put drop code Schedule to Workbook
    If Not Intersect(.Cells, Range("AI10")) Is Nothing Then
     If Range("AI10") <> "1" Then
        Worksheets("TANF Workbook").Range("AE45") = Range("AI10")
     Else
        Worksheets("TANF Workbook").Range("AE45") = ""
     End If
'Determining in Certification or Recertification from Workbook to Schedule
 '   ElseIf Not Intersect(.Cells, Range("G23")) Is Nothing Then
 '       If Range("G23") = "Certification" Then
 '           thisws.Range("M50") = "1"
 '       ElseIf Range("G23") = "Recertification" Then
 '           thisws.Range("M50") = "2"
 '       Else
 '           thisws.Range("M50") = ""
 '       End If
        
 'Determining Receipt of Expedited Service from Workbook to Schedule
 '    ElseIf Not Intersect(.Cells, Range("G28")) Is Nothing Then
 '       If Range("G28") = "Yes, Timely" Then
 '           thisws.Range("L55") = "1"
 '        ElseIf Range("G28") = "Yes, Untimely" Then
 '           thisws.Range("L55") = "2"
 '        ElseIf Range("G28") = "No" Then
 '           thisws.Range("L55") = "3"
 '        Else
 '           thisws.Range("L55") = ""
 '       End If
   
        
End If
    
End With

'Element code entry - create comments and data validation
If (Target.Row = 61 And Target.Column = 10) Or (Target.Row = 63 And Target.Column = 10) Or _
    (Target.Row = 65 And Target.Column = 10) Or (Target.Row = 67 And Target.Column = 10) Then
   
    If IsEmpty(Target.Value) Then
    Else
    Application.EnableEvents = False   'should be part of Change macro
    'Find start and stop rows with entered Element code
    startrow = 0
    endrow = 0
    lastrow& = Range("BP" & Rows.Count).End(xlUp).Row
    For i = 2 To lastrow
        If Target.Value = Range("BP" & i) And startrow = 0 Then
            startrow = i
        ElseIf Target.Value <> Range("BP" & i) And startrow <> 0 Then
            endrow = i - 1
            Exit For
        End If
    Next i
    'MsgBox startrow & " " & endrow
    
    'Create comment string
    TempStr = ""
    For j = startrow To endrow
        TempStr = TempStr & Range("BR" & j) & vbCrLf
    Next j
    Cells(Target.Row, 15).Comment.Text TempStr
    
    'Create data validation string
    TempStr = "=$BQ$" & startrow & ":$BQ$" & endrow
    Cells(Target.Row, 15).Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=TempStr
        .IgnoreBlank = True
        .InCellDropdown = False
        .InputTitle = ""
        .ErrorTitle = "Nature"
        .InputMessage = ""
        .ErrorMessage = _
        "Please enter a valid Nature code.  Refer to comment for description of codes."
        .ShowInput = True
        .ShowError = True
    End With
    End If 'blank target values

End If

 
Application.EnableEvents = True   'should be part of Change macro



End Sub

