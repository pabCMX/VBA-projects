Attribute VB_Name = "Module4"
Sub Find_Names()
' Loop through all names in workbook.
   For Each n In ActiveWorkbook.Names
      ' Check to see if the name refers to the ActiveSheet.
      If InStr(1, n.RefersTo, ActiveSheet.Name, vbTextCompare) > 0 Then
         ' If name refers to ActiveSheet, then find the intersection of the
         ' named range and the ActiveCell.
         If n.Name = "COLUMNS" Then
            MsgBox n.Name & " " & n.RefersTo
         End If
         'Set y = Intersect(ActiveCell, Range(n.RefersTo))
         ' Display a message box if the ActiveCell is in the named range.
         'If Not y Is Nothing Then MsgBox "Cell is in : " & n.Name
      End If
   Next
   MsgBox "No More Names!"
   ' Display message when finished.
End Sub

Sub group_box_outline_remove()
    ActiveSheet.GroupBoxes.Visible = False
End Sub
