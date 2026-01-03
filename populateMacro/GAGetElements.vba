Attribute VB_Name = "GAGetElements"
Sub GAgetresults()
    Select Case Left(Range("F201"), 3)
        Case "Two"
        
            'copy and paste words and borders
            Range("CaseTwo").Copy
            
            With Range("CaseTwo")
            Range("F204").Resize(.Rows.Count, .Columns.Count).Value = .Value
            End With

            'copy and paste buttons
            ActiveSheet.Shapes("CB 672").Copy
            Range("F205").Select
            ActiveSheet.Paste
            ActiveSheet.Shapes("CB 673").Copy
            Range("F206").Select
            ActiveSheet.Paste
            ActiveSheet.Shapes("CB 674").Copy
            Range("F209").Select
            ActiveSheet.Paste
            ActiveSheet.Shapes("CB 675").Copy
            Range("F211").Select
            ActiveSheet.Paste
            ActiveSheet.Shapes("CB 676").Copy
            Range("F213").Select
            ActiveSheet.Paste
            
'Non-parental case information for column 2

        Case "Non"
           'copy and paste words and borders
            Range("CaseNon").Copy
           
            
            With Range("CaseNon")
            Range("F204").Resize(.Rows.Count, .Columns.Count).Value = .Value
            End With

            'copy and paste buttons
            ActiveSheet.Shapes("CB 677").Copy
            Range("f204").Select
            ActiveSheet.Paste
            ActiveSheet.Shapes("CB 678").Copy
            Range("F206").Select
            ActiveSheet.Paste
            ActiveSheet.Shapes("CB 679").Copy
            Range("F209").Select
            ActiveSheet.Paste
            ActiveSheet.Shapes("CB 680").Copy
            Range("F211").Select
            ActiveSheet.Paste
            
            
'Drug for Case Findings for Column 2

        Case "Dru"
             'copy and paste words and borders
            Range("CaseDru").Copy

            With Range("CaseDru")
            Range("G204").Resize(.Rows.Count, .Columns.Count).Value = .Value
            End With

            'copy and paste buttons
            ActiveSheet.Shapes("CB 681").Copy
            Range("f204").Select
            ActiveSheet.Paste
            ActiveSheet.Shapes("CB 684").Copy
            Range("F206").Select
            ActiveSheet.Paste
            ActiveSheet.Shapes("CB 685").Copy
            Range("f209").Select
            ActiveSheet.Paste
            ActiveSheet.Shapes("CB 686").Copy
            Range("F211").Select
            ActiveSheet.Paste
            
            Range("BN515:BW515").Select
            Selection.Copy
            Range("G205").Select
            ActiveSheet.Paste
            
    Range("M210:P210").Select
         With Selection
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlBottom
         End With
            Selection.Merge
         Selection.Font.Bold = True
         Selection.NumberFormat = "@"
         With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
         End With

     Range("K204:M204").Select
         With Selection
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlBottom
         End With
            Selection.Merge
         Selection.Font.Bold = True
         Selection.NumberFormat = "@"
         With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
         End With
         
      Range("M207:O207").Select
         With Selection
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlBottom
         End With
            Selection.Merge
         Selection.Font.Bold = True
         Selection.NumberFormat = "@"
         With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
         End With
'Victim of Domestic Violence for Case information in Column 2

        Case "Vic"
             'copy and paste words and borders
            Range("CaseVic").Copy
            
            With Range("CaseVic")
            Range("G204").Resize(.Rows.Count, .Columns.Count).Value = .Value
            End With

            'copy and paste buttons
            ActiveSheet.Shapes("CB 843").Copy
            Range("f205").Select
            ActiveSheet.Paste
            ActiveSheet.Shapes("CB 861").Copy
            Range("F206").Select
            ActiveSheet.Paste
            
            
            Range("BN532:BV532").Select
            Selection.Copy
            Range("G205").Select
            ActiveSheet.Paste
            
            Range("M207:N207").Select
         With Selection
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlBottom
         End With
            Selection.Merge
         Selection.Font.Bold = True
         Selection.NumberFormat = "@"
         With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
         End With
  
            
   
 'Unemployable Permanent for Case Findings Column 2
         
        Case "Per"
            Range("CasePer").Copy
           
            With Range("CasePer")
            Range("F204").Resize(.Rows.Count, .Columns.Count).Value = .Value
            End With

            'copy and paste buttons
            ActiveSheet.Shapes("CB 848").Copy
            Range("f204").Select
            ActiveSheet.Paste
            ActiveSheet.Shapes("CB 849").Copy
            Range("F208").Select
            ActiveSheet.Paste
            ActiveSheet.Shapes("CB 850").Copy
            Range("F209").Select
            ActiveSheet.Paste
            
            Range("I210:K210").Select
         With Selection
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlBottom
         End With
            Selection.Merge
         Selection.Font.Bold = True
         Selection.NumberFormat = "@"
         With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
         End With
         
     Range("I206:L206").Select
         With Selection
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlBottom
         End With
            Selection.Merge
         Selection.Font.Bold = True
         Selection.NumberFormat = "@"
         With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
         End With
         
 'Unemployable Temporary for Case Findings Column 2
         
        Case "Tem"
            Range("CaseTem").Copy
        
            With Range("CaseTem")
            Range("F204").Resize(.Rows.Count, .Columns.Count).Value = .Value
            End With
            
            'copy and paste buttons
            ActiveSheet.Shapes("CB 852").Copy
            Range("f204").Select
            ActiveSheet.Paste
            ActiveSheet.Shapes("CB 853").Copy
            Range("F207").Select
            ActiveSheet.Paste
            ActiveSheet.Shapes("CB 962").Copy
            Range("F209").Select
            ActiveSheet.Paste
            
    Range("I210:K210").Select
         With Selection
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlBottom
         End With
            Selection.Merge
         Selection.Font.Bold = True
         Selection.NumberFormat = "@"
         With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
         End With
         
     Range("I206:L206").Select
         With Selection
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlBottom
         End With
            Selection.Merge
         Selection.Font.Bold = True
         Selection.NumberFormat = "@"
         With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
         End With
            
    End Select
    
    Select Case Left(Range("U201"), 3)
        Case "Two"
        
            'copy and paste words and borders
            Range("FindingsTwo").Copy
            'Range("F204").PasteSpecial
            
            With Range("FindingsTwo")
            Range("U204").Resize(.Rows.Count, .Columns.Count).Value = .Value
            End With

            'copy and paste only words - no borders
            'Range("BM92:BX100").Copy
            'Range("G205").PasteSpecial xlPasteValues
            
            'copy and paste buttons
            ActiveSheet.Shapes("CB 654").Copy
            Range("U204").Select
            ActiveSheet.Paste
            ActiveSheet.Shapes("CB 639").Copy
            Range("U206").Select
            ActiveSheet.Paste
            ActiveSheet.Shapes("CB 656").Copy
            Range("U210").Select
            ActiveSheet.Paste
            ActiveSheet.Shapes("CB 657").Copy
            Range("U212").Select
            ActiveSheet.Paste
            ActiveSheet.Shapes("CB 642").Copy
            Range("U217").Select
            ActiveSheet.Paste
            ActiveSheet.Shapes("CB 659").Copy
            Range("U218").Select
            ActiveSheet.Paste
            ActiveSheet.Shapes("CB 652").Copy
            Range("U223").Select
            ActiveSheet.Paste
            ActiveSheet.Shapes("TB 690").Copy
            Range("V208").Select
            ActiveSheet.Paste
            ActiveSheet.Shapes("TB 869").Copy
            Range("V213").Select
            ActiveSheet.Paste
            ActiveSheet.Shapes("TB 870").Copy
            Range("V219").Select
            ActiveSheet.Paste
            
        Case "Non"
         
            'copy and paste words and borders
            Range("FindingsNon").Copy
            
            With Range("FindingsNon")
            Range("U204").Resize(.Rows.Count, .Columns.Count).Value = .Value
            End With
   
            'copy and paste buttons
            ActiveSheet.Shapes("CB 661").Copy
            Range("U204").Select
            ActiveSheet.Paste
            ActiveSheet.Shapes("CB 653").Copy
            Range("U206").Select
            ActiveSheet.Paste
            ActiveSheet.Shapes("CB 663").Copy
            Range("U211").Select
            ActiveSheet.Paste
            ActiveSheet.Shapes("CB 664").Copy
            Range("U214").Select
            ActiveSheet.Paste
            ActiveSheet.Shapes("TB 868").Copy
            Range("V208").Select
            ActiveSheet.Paste
            Range("A204").Select
            
        Case "Dru"
            Range("FindingsDru").Copy
             With Range("FindingsDru")
            Range("V204").Resize(.Rows.Count, .Columns.Count).Value = .Value
            End With
   
            'copy and paste buttons
            ActiveSheet.Shapes("CB 665").Copy
            Range("U204").Select
            ActiveSheet.Paste
            ActiveSheet.Shapes("CB 963").Copy
            Range("U207").Select
            ActiveSheet.Paste
            ActiveSheet.Shapes("CB 668").Copy
            Range("U212").Select
            ActiveSheet.Paste
            ActiveSheet.Shapes("TB 964").Copy
            Range("V209").Select
            ActiveSheet.Paste
            ActiveSheet.Shapes("TB 965").Copy
            Range("V214").Select
            ActiveSheet.Paste
            ActiveSheet.Shapes("CB 1000").Copy
            Range("U216").Select
            ActiveSheet.Paste
            ActiveSheet.Shapes("CB 670").Copy
            Range("U218").Select
            ActiveSheet.Paste
             
            Range("CG515:CJ515").Select
            Selection.Copy
            Range("AA205").Select
            ActiveSheet.Paste
            
        Case "Vic"
            Range("FindingsVic").Copy
            With Range("FindingsVic")
            Range("V204").Resize(.Rows.Count, .Columns.Count).Value = .Value
            End With
   
            'copy and paste buttons
            ActiveSheet.Shapes("CB 829").Copy
            Range("U204").Select
            ActiveSheet.Paste
            ActiveSheet.Shapes("CB 830").Copy
            Range("U206").Select
            ActiveSheet.Paste
            ActiveSheet.Shapes("CB 5000").Copy
            Range("R209").Select
            ActiveSheet.Paste
            ActiveSheet.Shapes("CB 825").Copy
            Range("U211").Select
            ActiveSheet.Paste
            
             
            Range("CG532:CL532").Select
            Selection.Copy
            Range("AA204").Select
            ActiveSheet.Paste
            
        Range("AB211:Ae211").Select
         With Selection
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlBottom
         End With
            Selection.Merge
         Selection.Font.Bold = True
         Selection.NumberFormat = "@"
         With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
         End With
            
        Case "Per"
            Range("FindingsPer").Copy
            
               
            With Range("FindingsPer")
            Range("U204").Resize(.Rows.Count, .Columns.Count).Value = .Value
            End With
   
            'copy and paste buttons
            ActiveSheet.Shapes("CB 833").Copy
            Range("U204").Select
            ActiveSheet.Paste
            ActiveSheet.Shapes("CB 828").Copy
            Range("U206").Select
            ActiveSheet.Paste
            ActiveSheet.Shapes("CB 864").Copy
            Range("U210").Select
            ActiveSheet.Paste
            ActiveSheet.Shapes("CB 865").Copy
            Range("U212").Select
            ActiveSheet.Paste
            ActiveSheet.Shapes("CB 961").Copy
            Range("U214").Select
            ActiveSheet.Paste
            ActiveSheet.Shapes("TB 863").Copy
            Range("V208").Select
            ActiveSheet.Paste
            
         Range("AB211:AF211").Select
         With Selection
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlBottom
         End With
            Selection.Merge
         Selection.Font.Bold = True
         Selection.NumberFormat = "@"
         With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
         End With
         
         Range("AG215:AH215").Select
         With Selection
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlBottom
         End With
            Selection.Merge
         Selection.Font.Bold = True
         Selection.NumberFormat = "@"
         With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
         End With
         
          Range("AD215:AE215").Select
         With Selection
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlBottom
         End With
            Selection.Merge
         Selection.Font.Bold = True
         Selection.NumberFormat = "@"
         With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
         End With
            
            
        Case "Tem"
            Range("FindingsTem").Copy
            
            With Range("FindingsTem")
            Range("U204").Resize(.Rows.Count, .Columns.Count).Value = .Value
            End With
   
            'copy and paste buttons
            ActiveSheet.Shapes("CB 836").Copy
            Range("U204").Select
            ActiveSheet.Paste
            ActiveSheet.Shapes("CB 838").Copy
            Range("U206").Select
            ActiveSheet.Paste
            ActiveSheet.Shapes("CB 839").Copy
            Range("U211").Select
            ActiveSheet.Paste
            ActiveSheet.Shapes("CB 840").Copy
            Range("U213").Select
            ActiveSheet.Paste
            ActiveSheet.Shapes("TB 862").Copy
            Range("V208").Select
            ActiveSheet.Paste
            
         Range("AB212:AE212").Select
         With Selection
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlBottom
         End With
            Selection.Merge
         Selection.Font.Bold = True
         Selection.NumberFormat = "@"
         With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
         End With
    End Select
    
End Sub

Sub Clearleft()

Dim Sh As Shape
    Worksheets("GA Workbook").Range("f204:s225").UnMerge
    Worksheets("GA Workbook").Range("f204:S225").Clear
    
With Worksheets("GA Workbook")
   For Each Sh In .Shapes
       If Not Left(Sh.Name, 4) = "Drop" Then
       If Not Application.Intersect(Sh.TopLeftCell, _
        Range("f204:S225")) Is Nothing Then
        'MsgBox Sh.Name
         Sh.Delete
       End If
       End If
    Next Sh
End With
End Sub

Sub Clearright()

Dim Sh As Shape
    Worksheets("GA Workbook").Range("u204:AH225").UnMerge
    Worksheets("GA Workbook").Range("u204:AH225").Clear
With Worksheets("GA Workbook")
   For Each Sh In .Shapes
        If Not Left(Sh.Name, 4) = "Drop" Then
      
       If Not Application.Intersect(Sh.TopLeftCell, _
        Range("t204:AH225")) Is Nothing Then
        'MsgBox Sh.Name
         Sh.Delete
       End If
       End If
    Next Sh
End With
End Sub





