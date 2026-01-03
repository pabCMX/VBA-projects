Attribute VB_Name = "ImportAllMacros"
' ============================================================================
' ImportAllMacros - Import All VBA Modules and UserForms from Directory
' ============================================================================
' WHAT THIS MACRO DOES:
'   Prompts you to select a directory, then scans it for all .vba (module)
'   and .frm (UserForm) files and imports them into the active Excel workbook
'   with their correct names.
'
' USAGE:
'   1. Open the Excel workbook where you want to import the macros
'   2. Import this module first (ImportAllMacros.vba) manually via:
'      - Developer tab > Visual Basic > File > Import File
'   3. Enable "Trust access to the VBA project object model" in:
'      - File > Options > Trust Center > Trust Center Settings > Macro Settings
'   4. Run the ImportAllMacros subroutine (Alt+F8, select ImportAllMacros)
'   5. Select the folder containing your .vba and .frm files
'   6. All modules and forms will be imported with their original names
'
' FEATURES:
'   - Folder picker dialog to select source directory
'   - Automatically detects .vba (modules) and .frm (UserForms) files
'   - Removes existing components with the same name before importing
'   - Handles UserForm .frx files (binary data) automatically
'   - Provides progress feedback and summary report
'   - Skips this module itself to avoid re-importing
'   - Error handling with option to continue on errors
'
' IMPORTANT:
'   This macro requires "Trust access to the VBA project object model" to be enabled
'   in Excel's Trust Center settings, otherwise it will fail with a permission error.
'
' ============================================================================

Option Explicit

Sub ImportAllMacros()
    ' Main subroutine that prompts user to select a directory and imports all macros
    
    Dim sourcePath As String
    Dim folderDialog As Object ' FileDialog
    
    ' Create folder picker dialog
    Set folderDialog = Application.FileDialog(msoFileDialogFolderPicker)
    
    With folderDialog
        .Title = "Select folder containing .vba and .frm files"
        ' Try to set initial path to workbook location if available
        If Len(ThisWorkbook.Path) > 0 Then
            .InitialFileName = ThisWorkbook.Path
        End If
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            sourcePath = .SelectedItems(1)
        Else
            MsgBox "Import cancelled.", vbInformation
            Exit Sub
        End If
    End With
    
    ' Ensure path ends with backslash
    If Right(sourcePath, 1) <> "\" Then
        sourcePath = sourcePath & "\"
    End If
    
    ' Call the helper function to do the actual import
    Call ImportAllMacrosFromPath(sourcePath)
    
    Set folderDialog = Nothing
End Sub

Private Sub ImportAllMacrosFromPath(sourcePath As String)
    ' Helper subroutine that imports from a specified path
    ' Called by ImportAllMacros after user selects directory
    
    Dim fileName As String
    Dim filePath As String
    Dim importCount As Integer
    Dim skipCount As Integer
    Dim errorCount As Integer
    Dim vbComp As Object ' VBComponent
    Dim compName As String
    Dim msg As String
    
    ' Verify directory exists
    If Dir(sourcePath, vbDirectory) = "" Then
        MsgBox "Directory not found: " & sourcePath, vbCritical, "Error"
        Exit Sub
    End If
    
    ' Initialize counters
    importCount = 0
    skipCount = 0
    errorCount = 0
    
    ' Disable screen updating and alerts for faster execution
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    On Error GoTo ErrorHandler
    
    ' Import all .vba files (modules)
    fileName = Dir(sourcePath & "*.vba")
    Do While fileName <> ""
        filePath = sourcePath & fileName
        
        ' Skip this module itself
        If fileName <> "ImportAllMacros.vba" Then
            ' Extract component name from file (remove .vba extension)
            compName = Left(fileName, Len(fileName) - 4)
            
            ' Remove existing component if it exists
            On Error Resume Next
            Set vbComp = ThisWorkbook.VBProject.VBComponents(compName)
            If Not vbComp Is Nothing Then
                ThisWorkbook.VBProject.VBComponents.Remove vbComp
            End If
            On Error GoTo ErrorHandler
            
            ' Import the module
            ThisWorkbook.VBProject.VBComponents.Import filePath
            importCount = importCount + 1
            Debug.Print "Imported: " & fileName & " -> " & compName
        Else
            skipCount = skipCount + 1
        End If
        
        fileName = Dir ' Get next file
    Loop
    
    ' Import all .frm files (UserForms)
    fileName = Dir(sourcePath & "*.frm")
    Do While fileName <> ""
        filePath = sourcePath & fileName
        
        ' Extract component name from file (remove .frm extension)
        compName = Left(fileName, Len(fileName) - 4)
        
        ' Remove existing component if it exists
        On Error Resume Next
        Set vbComp = ThisWorkbook.VBProject.VBComponents(compName)
        If Not vbComp Is Nothing Then
            ThisWorkbook.VBProject.VBComponents.Remove vbComp
        End If
        On Error GoTo ErrorHandler
        
        ' Import the UserForm
        ThisWorkbook.VBProject.VBComponents.Import filePath
        
        ' Check if corresponding .frx file exists
        Dim frxPath As String
        frxPath = sourcePath & compName & ".frx"
        If Dir(frxPath) <> "" Then
            Debug.Print "Found .frx file for: " & compName
        End If
        
        importCount = importCount + 1
        Debug.Print "Imported: " & fileName & " -> " & compName
        
        fileName = Dir ' Get next file
    Loop
    
    ' Re-enable screen updating and alerts
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    ' Show summary message
    msg = "Import Complete!" & vbCrLf & vbCrLf
    msg = msg & "Successfully imported: " & importCount & " component(s)" & vbCrLf
    If skipCount > 0 Then
        msg = msg & "Skipped: " & skipCount & " component(s)" & vbCrLf
    End If
    If errorCount > 0 Then
        msg = msg & "Errors: " & errorCount & " component(s)" & vbCrLf
    End If
    msg = msg & vbCrLf & "Source directory: " & sourcePath
    
    MsgBox msg, vbInformation, "Macro Import"
    
    ' Clean up
    Set vbComp = Nothing
    Exit Sub
    
ErrorHandler:
    ' Re-enable screen updating and alerts
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    errorCount = errorCount + 1
    msg = "Error importing " & fileName & ": " & Err.Description & vbCrLf & vbCrLf
    msg = msg & "Continue with remaining files?"
    
    Dim response As VbMsgBoxResult
    response = MsgBox(msg, vbYesNo + vbCritical, "Import Error")
    
    If response = vbYes Then
        Resume Next
    Else
        MsgBox "Import cancelled by user.", vbInformation
        Exit Sub
    End If
End Sub

