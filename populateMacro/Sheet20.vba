VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet20"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)



'David McRitchie, 2004-09-26, programming,  Case -- Entire Row
  '  http://www.mvps.org/dmcritchie/excel/event.htm#case
  
  Dim i As Long, j As Long
  Dim startrow As Long, endrow As Long, lastrow As Long
  'MsgBox Target.Row & " " & Target.Column
  
 Application.EnableEvents = True   'should be part of Change macro
 
'Clear schedule if case is a drop
If Target.Row = 34 And Target.Column = 32 Then
    If (Target.Value = 2 Or Target.Value = 3) Then
        Application.EnableEvents = False   'should be part of Change macro
  
        Range("f29") = "" 'J
        Range("q29") = "" 'K
        Range("af29") = "" 'L
        Range("j34") = "" 'M(a)
        Range("u34") = "" 'M(b)
        Range("d46") = "" 'Q
        Range("i46") = "" 'R
        Range("n46") = "" 'S
        Range("t46") = "" 'T
        
    End If
 
'If Element code is entered, create comment and data validation for just this Element code
ElseIf Target.Row = 41 Then
  If (Target.Column = 5 Or Target.Column = 11 Or Target.Column = 17) And Target.Value <> "" Then
    Application.EnableEvents = False   'should be part of Change macro
    startrow = 0
    endrow = 0
    lastrow& = Range("BF" & Rows.Count).End(xlUp).Row
    'Find rows with data for this Element code
    For i = 2 To lastrow + 1
        If Target.Value = Range("BF" & i) And startrow = 0 Then
            startrow = i
        ElseIf Target.Value <> Range("BF" & i) And startrow <> 0 Then
            endrow = i - 1
            Exit For
        End If
    Next i
    'MsgBox startrow & " " & endrow
    
    'Create comment
    TempStr = ""
    For j = startrow To endrow
        TempStr = TempStr & Range("BH" & j) & vbCrLf
    Next j
    Cells(41, Target.Column + 18).Comment.Text TempStr
    
    'Create data validation
    TempStr = "=$BG$" & startrow & ":$BG$" & endrow
    Cells(41, Target.Column + 18).Select
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

  End If

End If

    Application.EnableEvents = True   'should be part of Change macro

End Sub


