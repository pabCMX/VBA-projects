VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet18"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Worksheet_Change(ByVal Target As Range)

  Dim i As Integer, j As Integer
  
 Application.EnableEvents = True   'should be part of Change macro
    For i = 3 To 12
 If Target.Row = 64 And Target.Column = i Then
    If Target.Value = "" Then

        Application.EnableEvents = False   'should be part of Change macro
 
        Cells(67, i) = ""
    End If

 End If
Next i
    Application.EnableEvents = True   'should be part of Change macro

End Sub

