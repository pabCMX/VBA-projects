Attribute VB_Name = "Module5"
Sub ExportAllMacros()
    ' This script exports all VBA modules, class modules, and forms
    ' from the current workbook to the populateMacro folder
    
    Dim wb As Workbook
    Dim vbComp As Object ' VBComponent
    Dim exportPath As String
    Dim fileName As String
    Dim fileNum As Integer
    Dim compType As String
    Dim fileExt As String
    Dim successCount As Integer
    Dim errorCount As Integer
    
    ' Set reference to current workbook
    Set wb = ActiveWorkbook
    
    ' Build path to populateMacro folder (relative to workbook location)
    ' If workbook is not saved, use the workspace path
    If Len(wb.Path) > 0 Then
        exportPath = wb.Path & "\populateMacro\"
    Else
        ' Fallback: use the workspace path if workbook is unsaved
        exportPath = "C:\Users\PHLaptop\Documents\Programming\VBA-projects\populateMacro\"
    End If
    
    ' Create folder if it doesn't exist
    If Dir(exportPath, vbDirectory) = "" Then
        MkDir exportPath
    End If
    
    ' Initialize counters
    successCount = 0
    errorCount = 0
    
    ' Disable screen updating for faster execution
    Application.ScreenUpdating = False
    
    ' Loop through all VB components (modules, classes, forms)
    For Each vbComp In wb.VBProject.VBComponents
        ' Determine component type and file extension
        Select Case vbComp.Type
            Case 1 ' vbext_ct_StdModule - Standard Module
                compType = "Module"
                fileExt = ".vba"
            Case 2 ' vbext_ct_ClassModule - Class Module
                compType = "Class"
                fileExt = ".vba"
            Case 3 ' vbext_ct_MSForm - UserForm
                compType = "Form"
                fileExt = ".frm"
            Case 100 ' vbext_ct_Document - Document Module (Worksheet/Workbook)
                compType = "Document"
                fileExt = ".vba"
            Case Else
                compType = "Unknown"
                fileExt = ".vba"
        End Select
        
        ' Skip empty components or document modules if desired
        ' (Document modules are typically ThisWorkbook and Sheet modules)
        If vbComp.CodeModule.CountOfLines > 0 Then
            ' Build filename: ComponentName + extension
            fileName = exportPath & vbComp.Name & fileExt
            
            ' Check if file already exists and add number if needed
            fileNum = 1
            Dim originalFileName As String
            originalFileName = fileName
            Do While Dir(fileName) <> ""
                fileName = exportPath & vbComp.Name & "_" & fileNum & fileExt
                fileNum = fileNum + 1
            Loop
            
            ' Export the component
            On Error Resume Next
            vbComp.Export fileName
            If Err.Number = 0 Then
                successCount = successCount + 1
                Debug.Print "Exported: " & vbComp.Name & " (" & compType & ") -> " & fileName
            Else
                errorCount = errorCount + 1
                Debug.Print "Error exporting " & vbComp.Name & ": " & Err.Description
                Err.Clear
            End If
            On Error GoTo 0
        End If
    Next vbComp
    
    ' Re-enable screen updating
    Application.ScreenUpdating = True
    
    ' Show summary message
    Dim msg As String
    msg = "Export Complete!" & vbCrLf & vbCrLf
    msg = msg & "Successfully exported: " & successCount & " component(s)" & vbCrLf
    If errorCount > 0 Then
        msg = msg & "Errors: " & errorCount & " component(s)" & vbCrLf
    End If
    msg = msg & vbCrLf & "Files saved to: " & exportPath
    
    MsgBox msg, vbInformation, "Macro Export"
    
    ' Clean up
    Set vbComp = Nothing
    Set wb = Nothing
End Sub

