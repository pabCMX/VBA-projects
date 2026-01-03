VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet27"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Private Sub Worksheet_Change(ByVal Target As Range)

  Dim i As Integer, j As Integer
  
    Application.EnableEvents = True   'should be part of Change macro
    For i = 2 To 14
        If Target.Row = 66 And Target.Column = i Then
            If IsEmpty(Cells(Target.Row, Target.Column)) Then
                Application.EnableEvents = False   'should be part of Change macro
                Cells(68, i) = ""
            End If
        ElseIf Target.Row = 25 And Target.Column = i Then
            If Target.Value = 2 Then
                Application.EnableEvents = False   'should be part of Change macro
                Cells(24, i) = 1
            ElseIf Target.Value = 1 Then
                Application.EnableEvents = False   'should be part of Change macro
                Cells(26, i) = ""
            End If
        End If
    Next i
    Application.EnableEvents = True   'should be part of Change macro

End Sub


