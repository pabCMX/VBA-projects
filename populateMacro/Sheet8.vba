VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
  
 
  Application.EnableEvents = True   'should be part of Change macro
  ' Vehicle Box values
If (Target.Row = 62) And Target.Column = 16 Then
    If Target.Value = 1 Then Range("V62") = "-"
    If Target.Value <> 1 Then Range("V62") = ""
End If

Application.EnableEvents = True   'should be part of Change macro

End Sub

