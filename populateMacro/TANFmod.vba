Attribute VB_Name = "TANFmod"
Sub tanf()

For i = 3 To 14

If Cells(68, i) > 0 Then

    If Cells(66, i) < 101 Then
        Hundint = Application.WorksheetFunction.RoundDown((Cells(66, i) / 100) - 1, 0)
    Else
        Hundint = Application.WorksheetFunction.RoundDown((Cells(66, i) / 100), 0)
    End If

 If Cells(68, i) > 14 Then
 MsgBox ("Proration tables only go up to 14 days.  Please ask your Supervisor how to prorate beyond 14 days.")
 Else
 Application.EnableEvents = False
 Select Case Cells(68, i).Value
 
  Case 1
    Cells(70, i) = 6.6 * Hundint + Round(0.0000472884 + 0.065710554 * (Cells(66, i) - 100 * Hundint), 1)
  Case 2
    Cells(70, i) = 13.1 * Hundint + Round(-0.00017559 + 0.131430335 * (Cells(66, i) - 100 * Hundint), 1)
  Case 3
  If Cells(66, i) = 87.5 Then
    Cells(70, i) = 17.2
  Else
    Cells(70, i) = 19.7 * Hundint + Round(-0.00010194 + 0.197144297 * (Cells(66, i) - 100 * Hundint), 1)
  End If
  Case 4
  If Cells(66, i) = 5.9 Then
    Cells(70, i) = 1.5
  Else
    Cells(70, i) = 26.3 * Hundint + Round(-0.000130429 + 0.262858301 * (Cells(66, i) - 100 * Hundint), 1)
  End If
  Case 5
  If Cells(66, i) = 6.9 Then
    Cells(70, i) = 2.2
  Else
    Cells(70, i) = 32.9 * Hundint + Round(-0.001141619 + 0.328581779 * (Cells(66, i) - 100 * Hundint), 1)
  End If
  Case 6
  If Cells(66, i) = 100.3 Then
    Cells(70, i) = 39.6
  Else
    Cells(70, i) = 39.4 * Hundint + Round(-0.00018061 + 0.3942901 * (Cells(66, i) - 100 * Hundint), 1)
  End If
  Case 7
  If Cells(66, i) = 50.3 Then
    Cells(70, i) = 23.2
  Else
    Cells(70, i) = 46 * Hundint + Round(0.001345477 + 0.459995182 * (Cells(66, i) - 100 * Hundint), 1)
  End If
  Case 8
    Cells(70, i) = 52.6 * Hundint + Round(-0.0000486654 + 0.5257144 * (Cells(66, i) - 100 * Hundint), 1)
  Case 9
    Cells(70, i) = 59.1 * Hundint + Round(-0.000156025 + 0.591429951 * (Cells(66, i) - 100 * Hundint), 1)
  Case 10
    Cells(70, i) = 65.7 * Hundint + Round(0.000167077 + 0.657139017 * (Cells(66, i) - 100 * Hundint), 1)
  Case 11
    Cells(70, i) = 72.3 * Hundint + Round(-0.0000937807 + 0.722855619 * (Cells(66, i) - 100 * Hundint), 1)
  Case 12
  If Cells(66, i) = 46.6 Then
    Cells(70, i) = 36.8
  Else
    Cells(70, i) = 78.9 * Hundint + Round(0.0000808591 + 0.788573487 * (Cells(66, i) - 100 * Hundint), 1)
  End If
  Case 13
    Cells(70, i) = 85.4 * Hundint + Round(0.0000748837 + 0.85428116 * (Cells(66, i) - 100 * Hundint), 1)
  Case 14
    Cells(70, i) = 92 * Hundint + Round(0.000024456 + 0.91999952 * (Cells(66, i) - 100 * Hundint), 1)
  
  End Select
  End If
   Application.EnableEvents = True

  End If
  
  Next i
  
  End Sub
  
  

Sub TANFclear()

    'Cells(3, 8) = 0
   Ans = MsgBox("Are you sure that you want to clear the entire Computation sheet?", vbYesNo)
   If Ans = 6 Then
    
    Sheets("TANF Computation").Unprotect "QC"
    
    Range("A:N").Select
    'clearing any contents with numbers that were keyed in
    Selection.SpecialCells(xlCellTypeConstants, 1).Select
    Selection.ClearContents
  
    'clearing name in cell B1
    'Range("B1").Select
    'Selection.ClearContents
    
     'clearing county in cell C2
    'Range("C2").Select
    'Selection.ClearContents
    
     'clearing review number in cell E2
    'Range("E2").Select
    'Selection.ClearContents
    
    Range("A6") = "1.  line              / "
    Range("A7") = "2.  line              / "
    Range("A8") = "3.  line              / "
    Range("A15") = "7.  line              / "
    Range("A16") = "8.  line              / "
    Range("A17") = "9.  line              / "
    Range("A18") = "10.  line              / "
    Range("B78") = "Comments:"
    
     For i = 3 To 11
     
     TempStr = "=IF(R49C" & i & ">R50C" & i & ",0,IF(AND(R72C" & i & "=1,R69C" & "=" & """""" & "),R70C" & i & "*0.75,IF(AND(R72C" & i & "=1,R70C" & i & "=" & """""" & "),R69C" & i & "*0.75,IF(AND(R69C" & i & "=" & """""" & ",R70C" & i & "=" & """""" & ")," & """""" & ",IF(R69C" & i & "=" & """""" & ",R70C" & i & ",R69C" & i & ")))))"
        
            'TempStr = "=IF(AND(R67C" & i & "=" & """""" & ",R68C" & i & "=" & """""" & ")," & """""" & ",IF(R67C" & i & "=" & """""" & ",R68C" & i & ",R67C" & i & "))"
            Cells(71, i).FormulaR1C1 = TempStr
        
    Next i
    
         For i = 13 To 14
     
     TempStr = "=IF(R49C" & i & ">R50C" & i & ",0,IF(AND(R72C" & i & "=1,R69C" & "=" & """""" & "),R70C" & i & "*0.75,IF(AND(R72C" & i & "=1,R70C" & i & "=" & """""" & "),R69C" & i & "*0.75,IF(AND(R69C" & i & "=" & """""" & ",R70C" & i & "=" & """""" & ")," & """""" & ",IF(R69C" & i & "=" & """""" & ",R70C" & i & ",R69C" & i & ")))))"
        
            'TempStr = "=IF(AND(R67C" & i & "=" & """""" & ",R68C" & i & "=" & """""" & ")," & """""" & ",IF(R67C" & i & "=" & """""" & ",R68C" & i & ",R67C" & i & "))"
            Cells(71, i).FormulaR1C1 = TempStr
        
    Next i

        
    Sheets("TANF Computation").Protect "QC"
   
    Range("A1").Select
   End If
End Sub
Sub TANFCompute()

    Application.GoTo reference:=ActiveSheet.Range("I1"), Scroll:=True
    Range("M6").Select

End Sub

Sub show_form1()
    UserForm1.Show
End Sub

Sub redisplayform1()
    UserForm1.Show
End Sub

Sub TANFfinalresults()

    columnletter = Right(Range("AL77"), 1)
    rownumber = 71
    'If columnletter = "B" Then rownumber = 71
    TempStr = columnletter & rownumber
    Range(TempStr) = Range("M71") + Range("N71")
    Range(TempStr).Select

End Sub
Sub show_form2()
    UserForm2.Show
End Sub

Sub redisplayform2()
    UserForm2.Show
End Sub

Sub finaldetermination()

  clet = Right(Range("AL78"), 1)
  
  ' If row 69 doesn't have a formula then value was obtained
  ' from
  If Range(clet & "71").HasFormula Then
    Range(clet & "6:" & clet & "9").Copy
    Range("C6").PasteSpecial
    Range(clet & "11").Copy
    Range("C11").PasteSpecial
    Range(clet & "15:" & clet & "20").Copy
    Range("C15").PasteSpecial
    Range(clet & "22").Copy
    Range("C22").PasteSpecial
    Range(clet & "24:" & clet & "26").Copy
    Range("C24").PasteSpecial
    Range(clet & "30").Copy
    Range("C30").PasteSpecial
    Range(clet & "44").Copy
    Range("C44").PasteSpecial
    Range(clet & "46").Copy
    Range("C46").PasteSpecial
    Range(clet & "57").Copy
    Range("C57").PasteSpecial
    Range(clet & "60").Copy
    Range("C60").PasteSpecial
    Range(clet & "68:" & clet & "69").Copy
    Range("C68").PasteSpecial
    Range(clet & "74").Copy
    Range("C74").PasteSpecial
    Range(clet & "76").Copy
    Range("C76").PasteSpecial
    Range(clet & "78").Copy
    Range("C78").PasteSpecial
  Else
    Range("C71") = Range(clet & "71")
  End If
  
  Range("C76").Select

End Sub



