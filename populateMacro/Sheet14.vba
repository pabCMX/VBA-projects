VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Worksheet_Activate()

End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
  
 
  Application.EnableEvents = True   'should be part of Change macro
  ' Vehicle Box values
If (Target.Row = 62) And Target.Column = 16 Then
    If Target.Value = 1 Then Range("V62") = "-"
    If Target.Value <> 1 Then Range("V62") = ""
End If

'not entitled to expedited snap
'If (Target.Row = 55) And Target.Column = 12 Then
'    If Target.Value = 3 Then
'        Range("D159") = 3
'    Else
'        Range("D159") = ""
'    End If
'End If


'If (Target.Row = 50) And Target.Column = 13 Then
'    If Target.Value = 1 Then Range("B157") = ""
'End If

'If (Target.Row = 157) And Target.Column = 2 Then
' Application.EnableEvents = False
'    If Range("M50") = 1 Then
'    'MsgBox ("If type of action is Certification, then Face to Face should be blank.")
'    Range("B157") = "-"
'    End If
'End If

'If (Target.Row = 50) And Target.Column = 13 Then
'    If Target.Value = 1 Then Range("C157") = ""
'End If

'If (Target.Row = 157) And Target.Column = 3 Then
' Application.EnableEvents = False
'    If Range("M50") = 1 Then
'    Range("C157") = "4"
'    End If
'End If

'allotment adjustment changes
'If (Target.Row = 50) And Target.Column = 28 Then
'    If Target.value = 1 Then Range("AI50") = "-"
'End If

     'allotment adjustment changes
If (Target.Row = 50) And Target.Column = 28 Then
    If Range("AB50") = 1 Then
        Range("AI50") = "-"
    ElseIf Range("AB50") = 2 Or Range("AB50") = 3 Then
        Range("AI50") = " "
    End If
End If

   'If there is a disposition code and usage of SUA is 1 then Proration of SUA = "-"
If (Target.Row = 22) And Target.Column = 3 Then
    If Target.Value = 1 And Range("W82") = "1" Then
    Range("AA82") = "-"
    End If
End If

    'Useage of SUA changes
If (Target.Row = 82) And Target.Column = 23 Then
    If Target.Value = 1 And Range("C22") = "1" Then
    Range("AA82") = "-"
    ElseIf Range("W82") > 1 Then
    Range("AA82") = ""
    End If
End If

 'Useage of SUA changes
'If (Target.Row = 82) And Target.Column = 27 Then
   ' If Range("W82") = 1 And Range("C22") = "1" Then
   ' Range("AA82") = "-"
   ' End If
'End If


'putting timeliness code in depending on the most recent cert action and type of action

'If (Target.Row = 149) And Target.Column = 3 Then
'    If Range("M50") = 1 And Range("C50") >= 44835 And Target.Value <> "" Then
'        If (Target.Value = 1 Or Target.Value = 2) Then
'        Else
'        MsgBox "You entered an invalid code.  Value should be a 1 or 2."
'        Application.EnableEvents = False
'        Range("C149") = ""
'        End If
'    ElseIf Range("M50") = 1 And Range("C50") < 44835 Then
'        If Target.Value <> 3 Then
'            MsgBox "You entered an invalid code.  Value should be a 3."
'            Application.EnableEvents = False
'            Range("C149") = ""
'        End If
'    ElseIf Range("M50") = 2 And Target.Value <> "" Then
'            If Target.Value <> 3 Then
'                MsgBox "You entered an invalid code.  Value should be a 3."
'                Application.EnableEvents = False
'                Range("C149") = ""
'            End If
'    End If
'End If

'Element code entry - create comments and data validation
If (Target.Row = 29 And Target.Column = 2) Or (Target.Row = 31 And Target.Column = 2) Or _
    (Target.Row = 33 And Target.Column = 2) Or (Target.Row = 35 And Target.Column = 2) Or _
    (Target.Row = 37 And Target.Column = 2) Or (Target.Row = 39 And Target.Column = 2) Or _
    (Target.Row = 41 And Target.Column = 2) Or (Target.Row = 43 And Target.Column = 2) Then
    
    If IsEmpty(Target.Value) Then
    Else
    Application.EnableEvents = False   'should be part of Change macro
    'Find start and stop rows with entered Element code
    startrow = 0
    endrow = 0
    lastrow& = Range("BT" & Rows.Count).End(xlUp).Row
    For i = 2 To lastrow
        If Target.Value = Range("BT" & i) And startrow = 0 Then
            startrow = i
        ElseIf Target.Value <> Range("BT" & i) And startrow <> 0 Then
            endrow = i - 1
            Exit For
        End If
    Next i
    'MsgBox startrow & " " & endrow
    
    'Create comment string
    TempStr = ""
    For j = startrow To endrow
        TempStr = TempStr & Range("BV" & j) & vbCrLf
    Next j
    Cells(Target.Row, 7).Comment.Text TempStr
    
    'Create data validation string
    TempStr = "=$BU$" & startrow & ":$BU$" & endrow
    Cells(Target.Row, 7).Select
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




