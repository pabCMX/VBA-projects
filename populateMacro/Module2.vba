Attribute VB_Name = "Module2"
Sub CopyModule(SourceWB As Workbook, strModuleName As String, _
    TargetWB As Workbook)
' copies a module from one workbook to another
' example:
'  CopyModule Workbooks("Book1.xls"), "Module1", _
    Workbooks("Book2.xls")
Dim strFolder As String, strTempFile As String
    strFolder = SourceWB.Path
    If Len(strFolder) = 0 Then strFolder = CurDir
    strFolder = strFolder & "\"
    strTempFile = strFolder & "~tmpexport.bas"
    On Error Resume Next
    SourceWB.VBProject.VBComponents(strModuleName).Export strTempFile
    TargetWB.VBProject.VBComponents.Import strTempFile
    Kill strTempFile
    On Error GoTo 0
End Sub

Sub CopyForm(SourceWB As Workbook, strModuleName As String, _
    TargetWB As Workbook)
    
    Dim strFolder As String, strTempFile As String
    
    strFolder = SourceWB.Path
    If Len(strFolder) = 0 Then strFolder = CurDir
    strFolder = strFolder & "\"
    strTempFile = strFolder & "~tmpexport.frm"
    On Error Resume Next
    SourceWB.VBProject.VBComponents(strModuleName).Export strTempFile
    TargetWB.VBProject.VBComponents.Import strTempFile
    Kill strTempFile
    On Error GoTo 0

End Sub


