VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub shapes_outline()
'ActiveSheet.GroupBoxes.Visible = False

writerow = 11
    For Each shp In ActiveSheet.Shapes
        writerow = writerow + 1
        Range("j" & writerow) = shp.Name
        'If Left(shp.Name, 5) = "Group" Then
         '   shp.Delete
'Application.ScreenUpdating = True
            'shp.Fill.ForeColor.SchemeColor = 8
            'shp.Visible = False
            'shp.Select
            'MsgBox "Is this the one " & shp.Name
            'End
            'shp.Visible = True
        'End If
    Next
End Sub

Sub clear_buttons()
    For Each shp In ActiveSheet.Shapes
        If Left(shp.Name, 12) = "OptionButton" Then
            shp.OLEFormat.Object.Object.Value = False
        End If
    Next
End Sub
